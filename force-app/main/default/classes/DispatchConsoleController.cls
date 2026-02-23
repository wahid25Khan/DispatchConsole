/**
 * @description Controller for the Dispatch Console LWC component.
 *              Handles project/team data retrieval and action item scheduling.
 *              Supports both authenticated and guest user access.
 * @author TLG Development Team
 * @date 2026-01-30
 */
public without sharing class DispatchConsoleController {
  // =============================================================================
  // CONSTANTS
  // =============================================================================

  /** Field API names for TLG_Task__c scheduling fields (source of truth) */
  private static final String TASK_START_DATE = 'TLG_Start_Date__c';
  private static final String TASK_START_TIME = 'Start_Time__c';
  private static final String TASK_END_DATE = 'TLG_End_Date__c';
  private static final String TASK_END_TIME = 'End_Time__c';

  // =============================================================================
  // PRIVATE HELPER METHODS
  // =============================================================================

  /**
   * @description Pads a number to 2 digits with leading zero
   * @param n The number to pad
   * @return String padded to 2 characters
   */
  private static String pad2(Integer n) {
    return String.valueOf(n).leftPad(2, '0');
  }

  /**
   * @description Checks if a status value indicates completed/closed
   * @param statusValue The status string to check
   * @return Boolean true if status starts with 'complete' or 'closed'
   */
  private static Boolean isCompletedStatusValue(String statusValue) {
    if (String.isBlank(statusValue)) {
      return false;
    }
    String v = statusValue.trim().toLowerCase();
    return v.startsWith('complete') || v.startsWith('closed');
  }

  /**
   * @description Parses a time string from the client into a Time object
   * @param timeValue Time string in format "HH:mm" or "HH:mm:ss"
   * @return Time object or null if blank
   * @throws AuraHandledException if format is invalid
   */
  private static Time parseClientTime(String timeValue) {
    if (String.isBlank(timeValue)) {
      return null;
    }
    String v = timeValue.trim();

    // Strip timezone suffix and milliseconds if present
    if (v.endsWith('Z')) {
      v = v.substring(0, v.length() - 1);
    }
    if (v.contains('.')) {
      v = v.substring(0, v.indexOf('.'));
    }

    try {
      List<String> parts = v.split(':');
      if (parts.size() < 2) {
        throw new AuraHandledException('Invalid time format: ' + timeValue);
      }
      Integer hh = Integer.valueOf(parts[0]);
      Integer mm = Integer.valueOf(parts[1]);
      Integer ss = (parts.size() >= 3) ? Integer.valueOf(parts[2]) : 0;
      return Time.newInstance(hh, mm, ss, 0);
    } catch (Exception e) {
      throw new AuraHandledException('Invalid time value: ' + timeValue);
    }
  }

  /**
   * @description Builds an IN clause string from a set of Ids
   * @param ids Set of Salesforce Ids
   * @return Comma-separated string of quoted Ids for SOQL IN clause
   */
  private static String buildIdInClause(Set<Id> ids) {
    List<String> parts = new List<String>();
    if (ids == null)
      return '';
    for (Id i : ids) {
      parts.add('\'' + String.escapeSingleQuotes((String) i) + '\'');
    }
    return String.join(parts, ',');
  }

  // =============================================================================
  // PUBLIC AURA-ENABLED METHODS
  // =============================================================================

  /**
   * @description Retrieves initial data for the Dispatch Console (teams and projects)
   * @return Map containing 'teams', 'projects', and empty 'actionItems' array
   */
  @AuraEnabled(cacheable=true)
  public static Map<String, Object> getInitialData() {
    Map<String, Object> response = new Map<String, Object>();

    // Fetch all Teams with their Members
    // Use dynamic SOQL to avoid local compile-time dependency on custom SObject types.
    response.put(
      'teams',
      Database.query(
        'SELECT Id, Name, ' +
          '(SELECT Id, Name, Role__c, TLG_Member__c, ' +
          'TLG_Member__r.Name, TLG_Member__r.Designation__c, TLG_Member__r.Allowed_Overlap__c ' +
          ' FROM Team_Members__r WHERE TLG_Member__c != NULL) ' +
          'FROM TLG_Team__c ORDER BY Name'
      )
    );

    // Fetch all Projects (Opportunities)
    response.put(
      'projects',
      [
        SELECT Id, Name, Team__c, Account.Name
        FROM Opportunity
        ORDER BY Name
      ]
    );

    // Tasks loaded on-demand per project
    response.put('actionItems', new List<SObject>());
    return response;
  }

  /**
   * @description Loads all dispatch data for a specific project
   * @param projectId The Opportunity Id to load data for
   * @return Map containing 'project', 'team', 'teamName', 'targets', and 'actionItems'
   */
  @AuraEnabled(cacheable=false)
  public static Map<String, Object> getProjectDispatchData(Id projectId) {
    Map<String, Object> response = new Map<String, Object>();
    if (projectId == null) {
      response.put('project', null);
      response.put('team', null);
      response.put('teamName', null);
      response.put('targets', new List<Case>());
      response.put('actionItems', new List<SObject>());
      return response;
    }

    Opportunity project = [
      SELECT Id, Name, Team__c, Account.Name
      FROM Opportunity
      WHERE Id = :projectId
      LIMIT 1
    ];

    response.put('project', project);

    // Team for the project (read-only header + member lanes)
    if (project.Team__c != null) {
      // Return a single team record (not a List) for the LWC.
      List<TLG_Team__c> teams = [
        SELECT
          Id,
          Name,
          (
            SELECT
              Id,
              Name,
              Role__c,
              TLG_Member__c,
              TLG_Member__r.Name,
              TLG_Member__r.Designation__c,
              TLG_Member__r.Allowed_Overlap__c
            FROM Team_Members__r
            WHERE TLG_Member__c != NULL
          )
        FROM TLG_Team__c
        WHERE Id = :project.Team__c
        LIMIT 1
      ];
      TLG_Team__c team = teams.isEmpty() ? null : teams[0];
      response.put('team', team);
      response.put('teamName', team == null ? null : team.Name);
    } else {
      response.put('team', null);
      response.put('teamName', null);
    }

    // Action Items for this Project, including Target and Assigned To.
    // Static SOQL with bind variables avoids quoting/parsing errors (e.g., unexpected token '006').
    List<TLG_Task__c> actionItems = [
      SELECT
        Id,
        Name,
        TLG_Opportunity__c,
        TLG_Opportunity__r.Name,
        TLG_Case__c,
        TLG_Case__r.CaseNumber,
        TLG_Case__r.Subject,
        TLG_Case__r.Status,
        TLG_Assign_To__c,
        TLG_Assign_To__r.Name,
        TLG_Priority__c,
        TLG_Status__c,
        TLG_Team__c,
        Duration__c,
        TLG_Detail__c,
        Progress__c,
        TLG_Start_Date__c,
        Start_Time__c,
        TLG_End_Date__c,
        End_Time__c
      FROM TLG_Task__c
      WHERE TLG_Opportunity__c = :projectId AND TLG_Case__c != NULL
      ORDER BY TLG_Case__r.CaseNumber, Name
      LIMIT 2000
    ];
    response.put('actionItems', actionItems);

    // Targets (Cases)
    Set<Id> caseIds = new Set<Id>();
    for (SObject ai : actionItems) {
      Id cid = (Id) ai.get('TLG_Case__c');
      if (cid != null) {
        caseIds.add(cid);
      }
    }

    List<Case> targets = new List<Case>();
    if (!caseIds.isEmpty()) {
      targets = [
        SELECT Id, CaseNumber, Subject, Status
        FROM Case
        WHERE Id IN :caseIds
        ORDER BY CaseNumber
      ];
    }
    response.put('targets', targets);
    return response;
  }

  /**
   * @description Retrieves all tasks for a given team across all its projects
   * @param teamId The TLG_Team__c Id
   * @return List of TLG_Task__c records with related data
   */
  @AuraEnabled(cacheable=false)
  public static List<SObject> getTasksByTeam(Id teamId) {
    if (teamId == null)
      return new List<SObject>();

    List<Opportunity> opps = [
      SELECT Id
      FROM Opportunity
      WHERE Team__c = :teamId
      ORDER BY Name
      LIMIT 2000
    ];

    if (opps.isEmpty())
      return new List<SObject>();

    Set<Id> oppIds = new Set<Id>();
    for (Opportunity o : opps)
      oppIds.add(o.Id);

    return [
      SELECT Id, Name, TLG_Opportunity__c, TLG_Opportunity__r.Name,
             TLG_Case__c, TLG_Case__r.CaseNumber, TLG_Case__r.Subject, TLG_Case__r.Status,
             TLG_Assign_To__c, TLG_Assign_To__r.Name,
             TLG_Priority__c, TLG_Status__c, TLG_Team__c, TLG_Detail__c, Progress__c,
             TLG_Start_Date__c, Start_Time__c, TLG_End_Date__c, End_Time__c, Duration__c
      FROM TLG_Task__c
      WHERE TLG_Opportunity__c IN :oppIds
      ORDER BY TLG_Opportunity__r.Name, Name
      LIMIT 2000
    ];
  }

  /**
   * @description Searches projects by name or account name
   * @param searchTerm The search string to filter projects
   * @return List of matching Opportunity records (max 50)
   */
  @AuraEnabled(cacheable=true)
  public static List<Opportunity> searchProjects(String searchTerm) {
    if (String.isBlank(searchTerm)) {
      return [
        SELECT Id, Name, Account.Name, Team__c
        FROM Opportunity
        ORDER BY Name
        LIMIT 20
      ];
    }

    String query = '%' + searchTerm + '%';
    return [
      SELECT Id, Name, Account.Name, Team__c
      FROM Opportunity
      WHERE Name LIKE :query OR Account.Name LIKE :query
      ORDER BY Name
      LIMIT 50
    ];
  }

  /**
   * @description Updates the schedule and assignment for an action item
   * @param taskId The TLG_Task__c Id to update
   * @param assignedToId The Portal_Users__c Id to assign to
   * @param startDate Schedule start date
   * @param startTime Schedule start time (HH:mm:ss format)
   * @param endDate Schedule end date
   * @param endTime Schedule end time (HH:mm:ss format)
   * @param overrideOverlap Not used (overlaps allowed by design)
   * @return Map with 'success' boolean, 'error' message if failed, 'task' if successful
   */
  @AuraEnabled
  public static Map<String, Object> updateActionItemSchedule(
    Id taskId,
    Id assignedToId,
    Date startDate,
    String startTime,
    Date endDate,
    String endTime,
    Boolean overrideOverlap
  ) {
    Map<String, Object> response = new Map<String, Object>();

    try {
      if (taskId == null) {
        response.put('success', false);
        response.put('error', 'Missing action item id');
        return response;
      }

      Boolean allNull = (startDate == null &&
      startTime == null &&
      endDate == null &&
      endTime == null);
      Boolean anyNull = (startDate == null ||
      startTime == null ||
      endDate == null ||
      endTime == null);

      if (!allNull && anyNull) {
        response.put('success', false);
        response.put(
          'error',
          'Start Date/Time and End Date/Time must be fully specified'
        );
        return response;
      }

      List<TLG_Task__c> rows = [
        SELECT Id, TLG_Assign_To__c, TLG_Start_Date__c, Start_Time__c,
               TLG_End_Date__c, End_Time__c, TLG_Status__c
        FROM TLG_Task__c
        WHERE Id = :taskId
        LIMIT 1
      ];

      if (rows.isEmpty()) {
        response.put('success', false);
        response.put('error', 'Action item not found');
        return response;
      }

      TLG_Task__c task = rows[0];

      // Use existing assignment if none provided (backwards compatibility)
      Id effectiveAssignedToId = assignedToId != null
        ? assignedToId
        : task.TLG_Assign_To__c;

      if (effectiveAssignedToId == null) {
        response.put('success', false);
        response.put('error', 'Assigned To is required');
        return response;
      }

      // Completed action items are locked
      if (isCompletedStatusValue(task.TLG_Status__c)) {
        response.put('success', false);
        response.put(
          'error',
          'Completed action items are locked and cannot be rescheduled.'
        );
        return response;
      }

      task.TLG_Assign_To__c = effectiveAssignedToId;

      if (allNull) {
        // Clear schedule
        task.TLG_Start_Date__c = null;
        task.Start_Time__c = null;
        task.TLG_End_Date__c = null;
        task.End_Time__c = null;
      } else {
        Time parsedStartTime = parseClientTime(startTime);
        Time parsedEndTime = parseClientTime(endTime);

        if (parsedStartTime == null || parsedEndTime == null) {
          response.put('success', false);
          response.put(
            'error',
            'Start Time and End Time must be fully specified'
          );
          return response;
        }

        task.TLG_Start_Date__c = startDate;
        task.Start_Time__c = parsedStartTime;
        task.TLG_End_Date__c = endDate;
        task.End_Time__c = parsedEndTime;

        DateTime startDt = DateTime.newInstance(startDate, parsedStartTime);
        DateTime endDt = DateTime.newInstance(endDate, parsedEndTime);

        if (endDt <= startDt) {
          response.put('success', false);
          response.put('error', 'End date/time must be after start date/time');
          return response;
        }
      }

      update task;

      response.put('success', true);
      response.put('task', task);
      return response;
    } catch (Exception e) {
      System.debug(LoggingLevel.ERROR, 'Error updating action item schedule: ' + e.getMessage() + '\n' + e.getStackTraceString());
      response.put('success', false);
      response.put('error', 'An error occurred while updating the schedule. Please try again.');
      return response;
    }
  }
}