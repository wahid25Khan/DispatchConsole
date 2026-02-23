/**
 * @file dispatcherConsole.js
 * @description Dispatch Console LWC - A project-based task scheduling component
 *              with drag-and-drop timeline, backlog management, and team lanes.
 * @author TLG Development Team
 * @version 2026-01-30
 */

import getInitialData from "@salesforce/apex/DispatchConsoleController.getInitialData";
import getProjectDispatchData from "@salesforce/apex/DispatchConsoleController.getProjectDispatchData";
import searchProjects from "@salesforce/apex/DispatchConsoleController.searchProjects";
import updateActionItemSchedule from "@salesforce/apex/DispatchConsoleController.updateActionItemSchedule";
import TIME_ZONE from "@salesforce/i18n/timeZone";
import { api, LightningElement, track, wire } from "lwc";

// =============================================================================
// CONSTANTS
// =============================================================================

/** Timeline grid configuration */
const SLOT_MINUTES = 15;
const SLOTS_PER_DAY = (24 * 60) / SLOT_MINUTES; // 96 slots per day
const BASE_SLOT_WIDTH = 24;
const ZOOM_LEVELS = [0.75, 1, 1.25, 1.5];
const RESOURCE_ROW_HEIGHT = 68;
const LANE_VIRTUAL_BUFFER = 5;

/** Build identifier for cache busting */
const BUILD_STAMP = "2026-01-30.4";

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Pads a number to 2 digits with leading zero
 * @param {number} n - Number to pad
 * @returns {string} Padded string
 */
function pad2(n) {
  return String(n).padStart(2, "0");
}

/**
 * Parses a date string into year, month, day components
 * @param {string} dateStr - Date in YYYY-MM-DD format
 * @returns {Object|null} Object with year, month, day or null if invalid
 */
function parseDateStr(dateStr) {
  const parts = String(dateStr || "").split("-");
  if (parts.length !== 3) return null;
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  if (
    !Number.isFinite(year) ||
    !Number.isFinite(month) ||
    !Number.isFinite(day)
  ) {
    return null;
  }
  return { year, month, day };
}

/**
 * Converts a date string to UTC noon (avoids timezone display issues)
 * @param {string} dateStr - Date in YYYY-MM-DD format
 * @returns {Date|null} Date object or null if invalid
 */
function dateStrToUtcMidnight(dateStr) {
  const p = parseDateStr(dateStr);
  if (!p) return null;
  return new Date(Date.UTC(p.year, p.month - 1, p.day, 12, 0, 0));
}

/**
 * Adds days to a date string
 * @param {string} dateStr - Base date in YYYY-MM-DD format
 * @param {number} days - Number of days to add (can be negative)
 * @returns {string|null} New date string or null if invalid
 */
function addDaysToDateStr(dateStr, days) {
  const d = dateStrToUtcMidnight(dateStr);
  if (!d) return null;
  d.setUTCDate(d.getUTCDate() + Number(days || 0));
  return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
}

/**
 * Calculates the difference in days between two date strings
 * @param {string} dateStrA - First date
 * @param {string} dateStrB - Second date
 * @returns {number|null} Difference (A - B) in days or null if invalid
 */
function diffDays(dateStrA, dateStrB) {
  const a = dateStrToUtcMidnight(dateStrA);
  const b = dateStrToUtcMidnight(dateStrB);
  if (!a || !b) return null;
  const ms = a.getTime() - b.getTime();
  return Math.round(ms / (24 * 60 * 60 * 1000));
}

/**
 * Gets today's date in a specific timezone
 * @param {string} timeZone - IANA timezone string
 * @returns {string} Date in YYYY-MM-DD format
 */
function todayInTimeZoneDateStr(timeZone) {
  if (
    typeof Intl !== "undefined" &&
    typeof Intl.DateTimeFormat === "function"
  ) {
    const fmt = new Intl.DateTimeFormat("en-US", {
      timeZone,
      year: "numeric",
      month: "2-digit",
      day: "2-digit"
    });
    const parts = fmt.formatToParts(new Date());
    const y = parts.find((p) => p.type === "year")?.value;
    const m = parts.find((p) => p.type === "month")?.value;
    const d = parts.find((p) => p.type === "day")?.value;
    if (y && m && d) return `${y}-${m}-${d}`;
  }
  const now = new Date();
  return `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;
}

// =============================================================================
// COMPONENT CLASS
// =============================================================================

/**
 * Dispatch Console LWC Component
 * @extends LightningElement
 */
export default class DispatchConsole extends LightningElement {
  // ===========================================================================
  // PUBLIC API PROPERTIES
  // ===========================================================================
  
  /**
   * @description Enable read-only mode for guest users (disables scheduling actions)
   * @type {Boolean}
   */
  @api readOnlyMode = false;

  // ===========================================================================
  // TRACKED PROPERTIES - Data State
  // ===========================================================================
  @track isLoading = true;
  @track errorMessage = "";
  @track warningMessage = "";

  // --- Data ---
  @track allProjects = [];
  @track projectSearchResults = [];
  @track selectedOpportunity = null;

  // Project-scoped Dispatch Console data (single active project)
  @track activeProjectRecord = null; // Opportunity (UI label: Project)
  @track activeTeam = null; // TLG_Team__c (UI label: Team)
  @track activeTeamName = null; // Team.Name (for header display)
  @track teamLanes = []; // Portal Users (UI label: Members)
  @track targets = []; // Case (UI label: Targets)
  @track actionItems = []; // TLG_Task__c (UI label: Action Items)
  @track unscheduledByTarget = []; // [{ targetId, caseNumber, subject, status, actionItems: [] }]
  @track scheduledActionItems = []; // normalized scheduled items for timeline rendering

  // --- UI state ---
  @track theme = "light";

  // Timeline view controls
  @track timelineRange = "week"; // day | multiday | week
  @track zoomIndex = 1; // ZOOM_LEVELS[1] = 1

  // Optional global search (filters both backlog + scheduled)
  @track globalSearchQuery = "";

  // Drag hover feedback
  @track dragOverResourceId = null;
  @track dragOverSlotIndex = null;

  // Project dropdown
  @track isProjectDropdownOpen = false;
  @track projectDropdownSearchTerm = "";

  // Sidebar project filter
  @track selectedProjectId = "";
  @track sidebarSearchTerm = "";
  @track isProjectSearchActive = false;

  // Backlog filters
  @track isFilterOpen = false;
  @track backlogSearchQuery = "";
  @track priorityFilter = "All";
  @track durationFilter = "All";

  // Drawer
  @track selectedTask = null;

  // Scheduling Modal
  @track isSchedulingModalOpen = false;
  @track schedulingTask = null;
  @track scheduleAssignedToId = null;
  @track scheduleStartDate = null;
  @track scheduleStartTime = null;
  @track scheduleEndDate = null;
  @track scheduleEndTime = null;
  @track estimatedDuration = null;

  // Resize (silent save, 15-min snap)
  _isResizing = false;
  _resizeTaskId = null;
  _resizeStartClientX = 0;
  _resizeStartEndSlot = 0;
  _resizeStartTaskSnapshot = null;
  _boundMouseMove = null;
  _boundMouseUp = null;

  // Note: overlaps are always allowed; the UI shows a subtle indicator when relevant.

  // Lane virtualization (vertical)
  _ganttScrollEl = null;
  _boundGanttScroll = null;
  _pendingScrollFrame = false;
  @track laneVirtualStart = 0;
  @track laneVirtualEnd = 0;

  // Click-outside handler for dropdown
  _boundDocumentClick = null;

  // Help modal
  @track isHelpModalOpen = false;

  connectedCallback() {
    try {
      console.info("dispatcherConsole build", BUILD_STAMP);
    } catch {
      // no-op
    }
    this._boundMouseMove = this.handleWindowMouseMove.bind(this);
    this._boundMouseUp = this.handleWindowMouseUp.bind(this);
    this._boundDocumentClick = this.handleDocumentClick.bind(this);
    if (globalThis?.addEventListener) {
      globalThis.addEventListener("mousemove", this._boundMouseMove);
      globalThis.addEventListener("mouseup", this._boundMouseUp);
      globalThis.document?.addEventListener("click", this._boundDocumentClick);
    }

    this._boundGanttScroll = this.handleGanttScroll.bind(this);
  }

  renderedCallback() {
    // Attach scroll listener once to drive lane virtualization.
    if (!this._ganttScrollEl) {
      const el = this.template.querySelector(".fs-gantt-view");
      if (el) {
        this._ganttScrollEl = el;
        this._ganttScrollEl.addEventListener("scroll", this._boundGanttScroll);
        // Initial computation
        this.updateLaneVirtualization(true);
      }
    }
  }

  disconnectedCallback() {
    if (globalThis?.removeEventListener) {
      globalThis.removeEventListener("mousemove", this._boundMouseMove);
      globalThis.removeEventListener("mouseup", this._boundMouseUp);
      globalThis.document?.removeEventListener(
        "click",
        this._boundDocumentClick
      );
    }
    this._boundMouseMove = null;
    this._boundMouseUp = null;
    this._boundDocumentClick = null;

    try {
      if (this._ganttScrollEl && this._boundGanttScroll) {
        this._ganttScrollEl.removeEventListener(
          "scroll",
          this._boundGanttScroll
        );
      }
    } catch {
      // no-op
    }
    this._ganttScrollEl = null;
    this._boundGanttScroll = null;
  }

  // Timeline start date (CST-centric) in YYYY-MM-DD.
  @track activeDateStr = todayInTimeZoneDateStr(TIME_ZONE);

  @wire(getInitialData)
  wiredData({ error, data }) {
    this.isLoading = false;
    if (data) {
      this.errorMessage = "";
      this.allProjects = Array.isArray(data.projects) ? data.projects : [];
      this.projectSearchResults = [];
    } else if (error) {
      console.error("DispatchConsoleController.getInitialData failed", error);
      this.errorMessage = this.formatApexError(error);
    }
  }

  async loadInitialData() {
    try {
      this.isLoading = true;
      const data = await getInitialData();
      this.errorMessage = "";
      this.allProjects = Array.isArray(data?.projects) ? data.projects : [];
      this.projectSearchResults = [];
    } catch (err) {
      console.error(
        "DispatchConsoleController.getInitialData (imperative) failed",
        err
      );
      this.errorMessage = this.formatApexError(err);
    } finally {
      this.isLoading = false;
    }
  }

  /**
   * Loads dispatch data for a specific project
   * @param {string} projectId - The Opportunity Id
   */
  async loadDispatchDataForProject(projectId) {
    if (!projectId) return;
    try {
      this.isLoading = true;
      const result = await getProjectDispatchData({ projectId });

      this.activeProjectRecord = result?.project || null;
      this.activeTeam = result?.team || null;
      this.activeTeamName = result?.teamName || result?.team?.Name || null;
      this.targets = Array.isArray(result?.targets) ? result.targets : [];
      this.actionItems = Array.isArray(result?.actionItems)
        ? result.actionItems
        : [];

      this.teamLanes = this.buildTeamLanesFromTeam(this.activeTeam);
      this.updateLaneVirtualization(true);
      this.unscheduledByTarget = this.buildUnscheduledByTarget(
        this.targets,
        this.actionItems
      );

      const scheduled = this.buildScheduledActionItems(
        this.actionItems,
        this.teamLanes
      );
      this.scheduledActionItems = this.markOverlapsForIndicators(
        scheduled,
        this.teamLanes
      );

      this.isLoading = false;
    } catch (err) {
      console.error("Error loading project dispatch data", err);
      this.errorMessage = `${this.formatApexError(err)} (build ${BUILD_STAMP})`;
      this.isLoading = false;
    }
  }

  buildTeamLanesFromTeam(teamSObject) {
    // Apex may return a single team record or a one-item list; normalize here.
    const team = Array.isArray(teamSObject) ? teamSObject[0] : teamSObject;
    const members = team?.Team_Members__r;
    if (!Array.isArray(members) || members.length === 0) return [];

    // IMPORTANT: lanes are Members (Portal_Users__c) derived from the Team.
    return members
      .map((tm) => {
        const memberId = tm?.TLG_Member__c;
        const memberName = tm?.TLG_Member__r?.Name || tm?.Name || "Member";
        if (!memberId) return null;
        return {
          id: memberId,
          name: memberName,
          role: tm?.Role__c || tm?.TLG_Member__r?.Designation__c || "Member",
          allowedOverlap: Boolean(tm?.TLG_Member__r?.Allowed_Overlap__c),
          avatar: (memberName || "?").substring(0, 2).toUpperCase()
        };
      })
      .filter(Boolean);
  }

  getTaskStartDateValue(task) {
    return task?.TLG_Start_Date__c || task?.startDate || null;
  }

  getTaskEndDateValue(task) {
    return task?.TLG_End_Date__c || task?.endDate || null;
  }

  isActionItemFullyScheduled(task) {
    const sd = this.getTaskStartDateValue(task);
    const st = task?.Start_Time__c || task?.startTime;
    const ed = this.getTaskEndDateValue(task);
    const et = task?.End_Time__c || task?.endTime;
    return Boolean(sd && st && ed && et);
  }

  isActionItemUnscheduled(task) {
    const sd = this.getTaskStartDateValue(task);
    const st = task?.Start_Time__c || task?.startTime;
    const ed = this.getTaskEndDateValue(task);
    const et = task?.End_Time__c || task?.endTime;
    return !sd && !st && !ed && !et;
  }

  /**
   * Groups unscheduled action items by their target (Case)
   * @param {Array} targets - Array of Case records
   * @param {Array} actionItems - Array of TLG_Task__c records
   * @returns {Array} Grouped action items by target
   */
  buildUnscheduledByTarget(targets, actionItems) {
    const byId = {};
    (targets || []).forEach((t) => {
      byId[t.Id] = {
        targetId: t.Id,
        caseNumber: t.CaseNumber,
        subject: t.Subject,
        status: t.Status,
        actionItems: []
      };
    });

    (actionItems || []).forEach((ai) => {
      const targetId = ai?.TLG_Case__c;
      const isUnscheduled = this.isActionItemUnscheduled(ai);

      if (!targetId || !byId[targetId]) return;
      if (!isUnscheduled) return;

      byId[targetId].actionItems.push({
        id: ai.Id,
        title: ai.Name,
        priority: ai?.TLG_Priority__c || "Medium",
        status: ai?.TLG_Status__c,
        teamId: ai?.TLG_Team__c,
        assignedToId: ai?.TLG_Assign_To__c,
        assignedToName: ai?.TLG_Assign_To__r?.Name || "",
        duration: Number(ai?.Duration__c) || null,
        description: ai?.TLG_Detail__c || "",
        percentComplete: ai?.Progress__c ?? null,
        caseId: targetId,
        caseNumber: byId[targetId].caseNumber,
        caseSubject: byId[targetId].subject,
        location:
          byId[targetId].subject || `Target ${byId[targetId].caseNumber}`
      });
    });

    return Object.values(byId).filter((g) => (g.actionItems || []).length > 0);
  }

  /**
   * Builds array of scheduled action items with timeline positioning data
   * @param {Array} actionItems - Array of TLG_Task__c records
   * @param {Array} lanes - Array of team member lanes
   * @returns {Array} Scheduled items with slot calculations
   */
  buildScheduledActionItems(actionItems, lanes) {
    const laneIds = new Set((lanes || []).map((l) => l.id));

    const fullyScheduled = (actionItems || []).filter((ai) =>
      this.isActionItemFullyScheduled(ai)
    );

    return fullyScheduled
      .map((ai) => {
        const assignedToId = ai?.TLG_Assign_To__c;
        const isLocked = this.isCompletedStatus(ai?.TLG_Status__c);
        const startDate = this.getTaskStartDateValue(ai);
        const startTime = ai?.Start_Time__c;
        const endDate = this.getTaskEndDateValue(ai);
        const endTime = ai?.End_Time__c;

        const rawStartSlot = this.calculateGlobalSlotIndex(
          startDate,
          startTime
        );
        const rawEndSlot = this.calculateGlobalSlotIndex(endDate, endTime);

        return {
          id: ai.Id,
          title: ai.Name,
          priority: ai?.TLG_Priority__c || "Medium",
          status: ai?.TLG_Status__c,
          isLocked,
          isDraggable: !isLocked && !this.readOnlyMode,
          teamId: ai?.TLG_Team__c,
          assignedToId,
          assignedToName: ai?.TLG_Assign_To__r?.Name || "",
          duration: Number(ai?.Duration__c) || 60,
          description: ai?.TLG_Detail__c || "",
          percentComplete: ai?.Progress__c ?? null,
          caseId: ai?.TLG_Case__c,
          caseNumber: ai?.TLG_Case__r?.CaseNumber || "",
          caseSubject: ai?.TLG_Case__r?.Subject || "",
          location:
            ai?.TLG_Case__r?.Subject ||
            (ai?.TLG_Case__r?.CaseNumber
              ? `Target ${ai.TLG_Case__r.CaseNumber}`
              : ""),
          startDate,
          startTime,
          endDate,
          endTime,
          rawStartSlot,
          rawEndSlot,
          isOverlapping: false,
          showOverlapIndicator: false,
          resourceId: assignedToId
        };
      })
      .filter((t) => Boolean(t.assignedToId && laneIds.has(t.assignedToId)));
  }

  markOverlapsForIndicators(tasks, lanes) {
    const items = Array.isArray(tasks) ? tasks : [];
    if (!items.length) return [];

    const allowedByLane = new Map(
      (Array.isArray(lanes) ? lanes : []).map((l) => [
        l.id,
        Boolean(l.allowedOverlap)
      ])
    );

    // Group tasks by lane
    const byLane = new Map();
    for (const t of items) {
      if (!t?.resourceId) continue;
      if (!byLane.has(t.resourceId)) byLane.set(t.resourceId, []);
      byLane.get(t.resourceId).push(t);
    }

    const overlapIds = new Set();

    for (const laneTasks of byLane.values()) {
      const sortable = (laneTasks || [])
        .filter(
          (t) =>
            Number.isFinite(t.rawStartSlot) && Number.isFinite(t.rawEndSlot)
        )
        .map((t) => ({ ...t }))
        .sort((a, b) => a.rawStartSlot - b.rawStartSlot);

      // Simple sweep: mark any task that intersects its predecessor
      for (let i = 1; i < sortable.length; i++) {
        const prev = sortable[i - 1];
        const cur = sortable[i];
        if (
          cur.rawStartSlot < prev.rawEndSlot &&
          cur.rawEndSlot > prev.rawStartSlot
        ) {
          overlapIds.add(prev.id);
          overlapIds.add(cur.id);
        }
      }
    }

    return items.map((t) => {
      const isOverlapping = overlapIds.has(t.id);
      const allowedOverlap = allowedByLane.get(t.resourceId) === true;
      return {
        ...t,
        isOverlapping,
        showOverlapIndicator: Boolean(isOverlapping && !allowedOverlap)
      };
    });
  }

  getScheduledActionItemById(taskId) {
    return (
      (this.scheduledActionItems || []).find((t) => t.id === taskId) || null
    );
  }

  updateScheduledActionItem(taskId, patch) {
    const items = Array.isArray(this.scheduledActionItems)
      ? this.scheduledActionItems
      : [];
    const next = items.map((t) => {
      if (t.id !== taskId) return t;
      return { ...t, ...patch };
    });
    this.scheduledActionItems = next;

    if (this.selectedTask?.id === taskId) {
      this.selectedTask = {
        ...this.selectedTask,
        ...patch
      };
    }
  }

  isCompletedStatus(statusValue) {
    const s = String(statusValue || "")
      .trim()
      .toLowerCase();
    // Be tolerant of admin variations like "Closed - Completed".
    return s.startsWith("closed") || s.startsWith("complete");
  }

  formatApexError(error) {
    if (!error) return "Unknown error.";
    const body = error.body;
    if (Array.isArray(body)) {
      return (
        body
          .map((e) => e?.message)
          .filter(Boolean)
          .join(" | ") || "An error occurred."
      );
    }
    if (body && typeof body === "object" && body.message) {
      return body.message;
    }
    if (typeof error.message === "string") {
      return error.message;
    }
    return "An error occurred while loading Dispatch Console.";
  }

  // --- Getters for computed UI properties ---
  get containerClass() {
    return `app-container ${this.theme}-mode`;
  }

  get themeIcon() {
    return this.theme === "dark" ? "utility:sun" : "utility:moon";
  }

  get activeDateLabel() {
    const start = this.activeDateStr;
    if (!start) return "";
    const d = dateStrToUtcMidnight(start);
    if (!d) return start;
    if (
      typeof Intl !== "undefined" &&
      typeof Intl.DateTimeFormat === "function"
    ) {
      return new Intl.DateTimeFormat(undefined, {
        timeZone: TIME_ZONE,
        weekday: "short",
        month: "short",
        day: "2-digit",
        year: "numeric"
      }).format(d);
    }
    return start;
  }

  get activeDateValue() {
    return this.activeDateStr || "";
  }

  get daysInView() {
    const v = String(this.timelineRange || "").toLowerCase();
    if (v === "day") return 1;
    if (v === "multiday") return 3;
    return 7;
  }

  get totalSlots() {
    return SLOTS_PER_DAY * this.daysInView;
  }

  get dayHeaders() {
    const start = this.activeDateStr;
    if (!start) return [];
    const days = [];
    for (let i = 0; i < this.daysInView; i++) {
      const dateStr = addDaysToDateStr(start, i);
      let label = dateStr;
      const d = dateStrToUtcMidnight(dateStr);
      if (
        d &&
        typeof Intl !== "undefined" &&
        typeof Intl.DateTimeFormat === "function"
      ) {
        label = new Intl.DateTimeFormat(undefined, {
          timeZone: TIME_ZONE,
          weekday: "short",
          month: "short",
          day: "2-digit"
        }).format(d);
      }
      days.push({
        index: i,
        dateStr,
        label,
        widthStyle: `width: ${SLOTS_PER_DAY * this.slotWidth}px;`
      });
    }
    return days;
  }

  get activeProject() {
    // Project-scoped Dispatch Console (strict): lanes are derived from the active Project's Team.
    if (!this.selectedOpportunity || !Array.isArray(this.teamLanes)) {
      return { id: null, name: null, teamName: null, team: [] };
    }

    const lanes = this.getVirtualizedLanes(this.teamLanes);
    const scheduled = Array.isArray(this.scheduledActionItems)
      ? this.scheduledActionItems
      : [];

    const enrichedTeam = lanes.map((lane) => {
      const laneTasks = scheduled
        .filter((t) => t.resourceId === lane.id)
        .map((t) => ({
          ...t,
          isScheduled: true,
          // timeline adapter fields
          calculatedDuration: t.duration
        }))
        .map((t) => this.enrichScheduledTaskForResource(t, lane.name));

      return {
        id: lane.id,
        name: lane.name,
        avatar: lane.avatar,
        role: lane.role,
        allowedOverlap: Boolean(lane.allowedOverlap),
        isDragOver: lane.id === this.dragOverResourceId,
        rowClass: `fs-resource-row slds-grid ${lane.id === this.dragOverResourceId ? "is-dragover" : ""}`,
        tasks: laneTasks
      };
    });

    return {
      id: this.activeProjectRecord?.Id,
      name: this.activeProjectRecord?.Name,
      teamName: this.activeTeamName,
      team: enrichedTeam
    };
  }

  get laneVirtualPaddingStyle() {
    const total = Array.isArray(this.teamLanes) ? this.teamLanes.length : 0;
    if (!total) return "";

    const start = Math.max(0, Number(this.laneVirtualStart) || 0);
    const end = Math.max(start, Number(this.laneVirtualEnd) || 0);
    const top = start * RESOURCE_ROW_HEIGHT;
    const bottom = Math.max(0, total - end) * RESOURCE_ROW_HEIGHT;
    return `padding-top: ${top}px; padding-bottom: ${bottom}px;`;
  }

  getVirtualizedLanes(allLanes) {
    const lanes = Array.isArray(allLanes) ? allLanes : [];
    const total = lanes.length;
    if (!total) return [];

    // If we haven't measured yet, return all lanes (safe default).
    const end = Number(this.laneVirtualEnd);
    if (!Number.isFinite(end) || end <= 0) return lanes;

    const start = Math.max(0, Number(this.laneVirtualStart) || 0);
    const clampedEnd = Math.min(total, end);
    return lanes.slice(start, clampedEnd);
  }

  handleGanttScroll() {
    if (this._pendingScrollFrame) return;
    this._pendingScrollFrame = true;

    const raf = globalThis?.requestAnimationFrame;
    if (typeof raf === "function") {
      raf(() => {
        this._pendingScrollFrame = false;
        this.updateLaneVirtualization(false);
      });
    } else {
      this._pendingScrollFrame = false;
      this.updateLaneVirtualization(false);
    }
  }

  updateLaneVirtualization(force) {
    const lanes = Array.isArray(this.teamLanes) ? this.teamLanes : [];
    const total = lanes.length;
    if (!total) {
      if (force) {
        this.laneVirtualStart = 0;
        this.laneVirtualEnd = 0;
      }
      return;
    }

    const scroller = this._ganttScrollEl;
    if (!scroller) {
      if (force) {
        this.laneVirtualStart = 0;
        this.laneVirtualEnd = total;
      }
      return;
    }

    // The timeline header stack is sticky; measure its height to offset virtualization.
    const header = this.template.querySelector(".fs-timeline-header");
    const rowsTop = (header?.offsetTop || 0) + (header?.offsetHeight || 0);

    const scrollTop = Number(scroller.scrollTop) || 0;
    const viewportHeight = Number(scroller.clientHeight) || 0;
    const effectiveTop = Math.max(0, scrollTop - rowsTop);

    const start = Math.max(
      0,
      Math.floor(effectiveTop / RESOURCE_ROW_HEIGHT) - LANE_VIRTUAL_BUFFER
    );
    const visibleCount =
      Math.ceil(viewportHeight / RESOURCE_ROW_HEIGHT) + LANE_VIRTUAL_BUFFER * 2;
    const end = Math.min(total, start + Math.max(1, visibleCount));

    if (
      !force &&
      start === this.laneVirtualStart &&
      end === this.laneVirtualEnd
    ) {
      return;
    }

    this.laneVirtualStart = start;
    this.laneVirtualEnd = end;
  }

  enrichScheduledTaskForResource(task, memberName) {
    const base = { ...task, isScheduled: true };
    const enriched = this.enrichTaskForTimeline(base);
    const assignedToName = memberName || "Unassigned";

    return {
      ...enriched,
      assignedToName,
      ariaLabel: `${enriched.title || "Task"}. ${enriched.timeRange || ""}. Assigned to ${assignedToName}.`
    };
  }

  matchesDurationFilter(durationMinutes) {
    if (this.durationFilter === "All") return true;
    const d = Number(durationMinutes) || 0;
    if (this.durationFilter === "short") return d <= 60;
    if (this.durationFilter === "med") return d > 60 && d <= 120;
    if (this.durationFilter === "long") return d > 120;
    return true;
  }

  buildTargetLabel(task) {
    const caseNumber = task?.caseNumber;
    const caseSubject = task?.caseSubject;

    if (!caseNumber) return "";
    if (!caseSubject) return String(caseNumber);
    return `${caseNumber} - ${caseSubject}`;
  }

  get filteredUnscheduledByTarget() {
    const groups = Array.isArray(this.unscheduledByTarget)
      ? this.unscheduledByTarget
      : [];
    const q = (this.backlogSearchQuery || "").toLowerCase();
    const hasQuery = Boolean(q);
    const gq = (this.globalSearchQuery || "").toLowerCase();
    const hasGlobal = Boolean(gq);
    const priority = this.priorityFilter;

    return groups
      .map((g) => {
        // Show Subject as the target label, fallback to case number if no subject
        const displayLabel = g.subject || `Case ${g.caseNumber}` || "Target";

        const actionItems = (g.actionItems || [])
          .filter((ai) => {
            const title = String(ai.title || "").toLowerCase();
            const assigned = String(ai.assignedToName || "").toLowerCase();
            const targetLabel = String(ai.targetLabel || "").toLowerCase();
            const haystack = `${title} ${assigned} ${targetLabel}`;
            if (hasQuery && !haystack.includes(q)) return false;
            if (hasGlobal && !haystack.includes(gq)) return false;
            return true;
          })
          .filter((ai) => priority === "All" || ai.priority === priority)
          .map((ai) => {
            const isLocked = this.isCompletedStatus(ai.status);
            return {
              ...ai,
              isLocked,
              isDraggable: !isLocked && !this.readOnlyMode,
              priorityClass: `badge-priority ${(ai.priority || "Medium").toLowerCase()}`,
              statusClass: `dc-status-pill ${
                String(ai.status || "")
                  .toLowerCase()
                  .includes("complete") ||
                String(ai.status || "")
                  .toLowerCase()
                  .includes("closed")
                  ? "is-closed"
                  : ""
              }`
            };
          });

        return {
          ...g,
          displayLabel,
          actionItems
        };
      })
      .filter((g) => (g.actionItems || []).length > 0);
  }

  get unscheduledCount() {
    const groups = this.filteredUnscheduledByTarget;
    let n = 0;
    (groups || []).forEach((g) => (n += (g.actionItems || []).length));
    return n;
  }

  findUnscheduledActionItemById(actionItemId) {
    const groups = Array.isArray(this.unscheduledByTarget)
      ? this.unscheduledByTarget
      : [];
    for (const g of groups) {
      const ai = (g.actionItems || []).find((x) => x.id === actionItemId);
      if (ai) return ai;
    }
    return null;
  }

  get timeAxisTicks() {
    // Performance: avoid rendering a DOM node per 15-minute slot.
    // We render a single strip with a 15-minute grid background, plus hour (or 2-hour) labels.
    const ticks = [];
    const w = this.slotWidth;
    const hoursInView = this.daysInView * 24;

    // Reduce label density when zoomed out.
    const tickEveryHours = w < 16 ? 2 : 1;

    for (let h = 0; h < hoursInView; h += tickEveryHours) {
      const hourOfDay = h % 24;
      const dayOffset = Math.floor(h / 24);
      const slotIndex =
        dayOffset * SLOTS_PER_DAY + hourOfDay * (60 / SLOT_MINUTES);
      const left = slotIndex * w;

      const ampm = hourOfDay >= 12 ? "PM" : "AM";
      const h12 = hourOfDay % 12 || 12;
      const label = `${h12}${tickEveryHours >= 2 ? "" : ":00"} ${ampm}`;

      ticks.push({
        key: `d${dayOffset}h${hourOfDay}`,
        style: `left: ${left}px;`,
        label
      });
    }

    return ticks;
  }

  get slotsSurfaceStyle() {
    // Width for the scrollable slot area (excluding the sticky resource column).
    const w = this.slotWidth;
    return `width: ${this.totalSlots * w}px; min-width: ${this.totalSlots * w}px;`;
  }

  get timelineWidthStyle() {
    const w = this.slotWidth;
    return `--dc-slot-width: ${w}px; width: ${280 + this.totalSlots * w}px;`;
  }

  get zoomLabel() {
    const z = this.zoomFactor;
    return `${Math.round(z * 100)}%`;
  }

  get zoomFactor() {
    const idx = Number(this.zoomIndex);
    if (!Number.isInteger(idx) || idx < 0 || idx >= ZOOM_LEVELS.length)
      return 1;
    return ZOOM_LEVELS[idx];
  }

  get slotWidth() {
    return Math.max(12, Math.round(BASE_SLOT_WIDTH * this.zoomFactor));
  }

  get weekendBlocks() {
    const blocks = [];
    const startObj = dateStrToUtcMidnight(this.activeDateStr);
    if (!startObj) return [];

    for (let i = 0; i < this.daysInView; i++) {
      const currentDay = new Date(startObj);
      currentDay.setUTCDate(startObj.getUTCDate() + i);
      const dayOfWeek = currentDay.getUTCDay(); // 0(Sun)..6(Sat)

      if (dayOfWeek === 0 || dayOfWeek === 6) {
        const leftPx = i * SLOTS_PER_DAY * this.slotWidth;
        const widthPx = SLOTS_PER_DAY * this.slotWidth;
        blocks.push({
          key: `weekend-${i}`,
          style: `left: ${leftPx}px; width: ${widthPx}px;`
        });
      }
    }
    return blocks;
  }

  get currentTimeMarkerStyle() {
    const now = new Date();
    const todayStr = todayInTimeZoneDateStr(TIME_ZONE);
    const dDiff = diffDays(todayStr, this.activeDateStr);

    if (dDiff < 0 || dDiff >= this.daysInView) return "display: none;";

    const minutesFromMidnight = now.getHours() * 60 + now.getMinutes();
    const totalSlots =
      dDiff * SLOTS_PER_DAY + minutesFromMidnight / SLOT_MINUTES;

    const leftPx = totalSlots * this.slotWidth;
    return `left: ${leftPx}px;`;
  }

  get currentTimeMarkerContainerStyle() {
    const now = new Date();
    const todayStr = todayInTimeZoneDateStr(TIME_ZONE);
    const dDiff = diffDays(todayStr, this.activeDateStr);

    // Hide if current time is not within the visible date range
    if (dDiff < 0 || dDiff >= this.daysInView) return "display: none;";

    const minutesFromMidnight = now.getHours() * 60 + now.getMinutes();
    const totalSlots =
      dDiff * SLOTS_PER_DAY + minutesFromMidnight / SLOT_MINUTES;

    // Account for the resource column width (280px)
    const leftPx = 280 + totalSlots * this.slotWidth;
    return `left: ${leftPx}px;`;
  }

  get formattedEstimatedDuration() {
    const min = Number(this.estimatedDuration);
    if (!min || min <= 0) return "";
    const h = Math.floor(min / 60);
    const m = min % 60;
    if (h > 0) return `${h}h ${m}m (${min} min)`;
    return `${m} min`;
  }

  get dropIndicatorStyle() {
    if (!Number.isInteger(this.dragOverSlotIndex)) return "";
    const left = this.dragOverSlotIndex * this.slotWidth;
    return `left: ${left}px;`;
  }

  get dropIndicatorLabel() {
    if (!Number.isInteger(this.dragOverSlotIndex)) return "";
    const dt = this.getDateTimeFromSlot(this.dragOverSlotIndex);
    if (!dt?.date || !dt?.time) return "";
    return `${dt.time}`;
  }

  get filterBtnVariant() {
    return this.isFilterOpen ? "brand" : "bare";
  }

  get drawerClass() {
    return `detail-drawer ${this.selectedTask ? "open" : ""}`;
  }

  get isSelectedTaskScheduled() {
    const t = this.selectedTask;
    if (!t) return false;
    return Boolean(
      (t.assignedToId || t.resourceId) &&
      t.startDate &&
      t.startTime &&
      t.endDate &&
      t.endTime
    );
  }

  get selectedTaskProgressStyle() {
    const p = Number(this.selectedTask?.percentComplete) || 0;
    return `width: ${Math.min(100, Math.max(0, p))}%;`;
  }

  /**
   * @description Checks if scheduling actions are disabled (read-only mode for guest users)
   * @returns {Boolean} True if actions should be disabled
   */
  get isActionsDisabled() {
    return Boolean(this.readOnlyMode);
  }

  /**
   * @description Checks if drag-and-drop is allowed
   * @returns {Boolean} True if dragging is allowed
   */
  get isDragEnabled() {
    return !this.readOnlyMode;
  }

  get priorityOptions() {
    return [
      { label: "All Levels", value: "All" },
      { label: "High", value: "High" },
      { label: "Medium", value: "Medium" },
      { label: "Low", value: "Low" }
    ];
  }

  get durationOptions() {
    return [
      { label: "Any Duration", value: "All" },
      { label: "Short (≤ 1h)", value: "short" },
      { label: "Med (1-2h)", value: "med" },
      { label: "Long (2h+)", value: "long" }
    ];
  }

  get teamMemberOptions() {
    return (this.teamLanes || []).map((m) => ({ label: m.name, value: m.id }));
  }

  enrichTask(task) {
    const slotIndex = Number(task.slotIndex) || 0;
    const duration = Number(task.duration) || 60;
    const start = this.getSlotTime(slotIndex);
    const end = this.getSlotTime(
      slotIndex + Math.ceil(duration / SLOT_MINUTES)
    );
    const width = (duration / SLOT_MINUTES) * this.slotWidth;
    const left = slotIndex * this.slotWidth;

    const memberName =
      (this.teamLanes || []).find((m) => m.id === task.resourceId)?.name ||
      task.assignedToName ||
      "Unknown";
    const priority = task.priority || "Medium";

    return {
      ...task,
      duration,
      slotIndex,
      timeRange: `${start} - ${end}`,
      blockClass: `scheduled-block ${priority.toLowerCase()} ${task.showOverlapIndicator ? "has-overlap" : ""} ${task.isLocked ? "is-locked" : ""}`,
      blockStyle: `left: ${left + 6}px; width: ${Math.max(0, width - 12)}px;`,
      priorityClass: `badge-priority ${priority.toLowerCase()}`,
      assignedToName: memberName,
      showOverlapIndicator: Boolean(task.showOverlapIndicator),
      ariaLabel: `${task.title || "Task"}. ${start} to ${end}. Assigned to ${memberName}.`
    };
  }

  getSlotTime(i) {
    if (!Number.isFinite(i)) return "";
    const slotInDay =
      ((Math.floor(i) % SLOTS_PER_DAY) + SLOTS_PER_DAY) % SLOTS_PER_DAY;
    const minutesFromMidnight = slotInDay * SLOT_MINUTES;
    const hour = Math.floor(minutesFromMidnight / 60);
    const minute = minutesFromMidnight % 60;
    const ampm = hour >= 12 ? "PM" : "AM";
    const h12 = hour % 12 || 12;
    return `${h12}:${pad2(minute)} ${ampm}`;
  }

  // --- Event Handlers ---
  handleDateChange(e) {
    const dateValue = e.target.value;
    if (dateValue) {
      // Date picker returns YYYY-MM-DD; treat it as the start date for the CST-rendered timeline.
      this.activeDateStr = String(dateValue);
    }
  }

  toggleTheme() {
    this.theme = this.theme === "dark" ? "light" : "dark";
  }

  toggleFilterPanel() {
    this.isFilterOpen = !this.isFilterOpen;
  }

  handleBacklogSearch(e) {
    this.backlogSearchQuery = e.target.value;
  }

  handlePriorityChange(e) {
    this.priorityFilter = e.detail.value;
  }

  handleDurationChange(e) {
    this.durationFilter = e.detail.value;
  }

  stopPropagation(e) {
    e.stopPropagation();
  }

  findTaskById(taskId) {
    if (!taskId) return null;

    // Project-scoped mode (single active Project)
    const scheduled = (this.scheduledActionItems || []).find(
      (t) => t.id === taskId
    );
    if (scheduled) {
      return {
        ...scheduled,
        resourceId: scheduled.assignedToId
      };
    }

    const unscheduled = this.findUnscheduledActionItemById(taskId);
    if (unscheduled) return { ...unscheduled };

    return null;
  }

  openBacklogTask(e) {
    const id = e?.currentTarget?.dataset?.id;
    const task = this.findTaskById(id);
    if (!task) return;

    const priority = task.priority || "Medium";
    const duration = Number(task.duration) || 60;

    this.selectedTask = {
      ...task,
      duration,
      priority,
      priorityClass: `badge-priority ${priority.toLowerCase()}`,
      timeRange: task.timeRange || "Not scheduled",
      assignedToName: task.assignedToName || "Unassigned"
    };
  }

  openUnscheduledActionItem(e) {
    const id = e?.currentTarget?.dataset?.id;
    const ai = this.findUnscheduledActionItemById(id);
    if (!ai) return;

    const priority = ai.priority || "Medium";
    const duration = Number(ai.duration) || 0;

    this.selectedTask = {
      ...ai,
      id: ai.id,
      title: ai.title,
      duration,
      priority,
      priorityClass: `badge-priority ${priority.toLowerCase()}`,
      timeRange: "Not scheduled"
    };
  }

  handleBacklogKeyDown(e) {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      this.openBacklogTask(e);
    }
  }

  handleTaskBlockKeyDown(e) {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      this.openDrawer(e);
    }
  }

  handleTimelineSlotClick(e) {
    try {
      const slotIdx = Number(e?.currentTarget?.dataset?.slotIndex);
      if (!Number.isInteger(slotIdx)) return;
      const resourceId = e?.currentTarget?.dataset?.resourceId;

      // Resource view: schedule the currently selected task onto the chosen resource/slot.
      if (!resourceId) return;
      const selectedTaskId = this.selectedTask?.id;
      if (!selectedTaskId) {
        // Keep this lightweight: no UI spam, just guidance.
        // If you want a toast later, we can swap this.
        console.warn(
          "Select a task (from backlog) first, then click a time slot to schedule it."
        );
        return;
      }

      this.openSchedulingModal(this.selectedTask, slotIdx, resourceId, null);
    } catch (err) {
      console.error("Error handling timeline slot click", err);
    }
  }

  getSlotIndexFromPointerEvent(e) {
    const el = e?.currentTarget;
    if (!el?.getBoundingClientRect) return null;
    const rect = el.getBoundingClientRect();
    const x = (Number(e?.clientX) || 0) - rect.left;
    if (!Number.isFinite(x)) return null;
    const slotIdx = Math.floor(x / this.slotWidth);
    if (!Number.isFinite(slotIdx)) return null;
    return Math.min(this.totalSlots - 1, Math.max(0, slotIdx));
  }

  handleTimelineRowClick(e) {
    try {
      if (this._isResizing) return;

      const resourceId = e?.currentTarget?.dataset?.resourceId;
      if (!resourceId) return;

      const slotIdx = this.getSlotIndexFromPointerEvent(e);
      if (!Number.isInteger(slotIdx)) return;

      const selectedTaskId = this.selectedTask?.id;
      if (!selectedTaskId) {
        console.warn(
          "Select a task (from backlog) first, then click a time slot to schedule it."
        );
        return;
      }

      this.openSchedulingModal(this.selectedTask, slotIdx, resourceId, null);
    } catch (err) {
      console.error("Error handling timeline row click", err);
    }
  }

  handleTimelineRowDrop(e) {
    e.preventDefault();
    const resourceId = e?.currentTarget?.dataset?.resourceId;
    if (!resourceId) return;

    const slotIdx = this.getSlotIndexFromPointerEvent(e);
    if (!Number.isInteger(slotIdx)) return;

    const raw = e?.dataTransfer?.getData("task_payload");
    if (!raw) return;
    try {
      const payload = JSON.parse(raw);
      this.dragOverResourceId = null;
      this.dragOverSlotIndex = null;
      this.openSchedulingModal(payload, slotIdx, resourceId, null);
    } catch (err) {
      console.error("Invalid drag payload", err);
      this.errorMessage = "Unable to schedule: invalid drag payload.";
    }
  }

  handleTimelineRowDragOver(e) {
    e.preventDefault();
    const resourceId = e?.currentTarget?.dataset?.resourceId;
    if (!resourceId) return;

    const slotIdx = this.getSlotIndexFromPointerEvent(e);
    if (!Number.isInteger(slotIdx)) return;

    this.dragOverResourceId = resourceId;
    this.dragOverSlotIndex = slotIdx;

    if (e.dataTransfer) {
      e.dataTransfer.dropEffect = "move";
    }
  }

  handleTimelineRowDragLeave(e) {
    const resourceId = e?.currentTarget?.dataset?.resourceId;
    if (resourceId && resourceId === this.dragOverResourceId) {
      this.dragOverResourceId = null;
      this.dragOverSlotIndex = null;
    }
  }

  openDrawer(e) {
    try {
      // Prevent timeline surface click handler from firing when clicking a scheduled block.
      e?.stopPropagation?.();
    } catch {
      // no-op
    }
    const id = e.currentTarget.dataset.id;

    const task = this.findTaskById(id);
    if (!task) return;

    if (this.isActionItemFullyScheduled(task)) {
      this.selectedTask = this.enrichTaskForTimeline({
        ...task,
        isScheduled: true
      });
      return;
    }

    // Unscheduled action item: show basic details in the drawer
    this.openUnscheduledActionItem({ currentTarget: { dataset: { id } } });
  }

  closeDrawer() {
    this.selectedTask = null;
  }

  handleScheduleFromDrawer() {
    const task = this.selectedTask;
    if (!task || task.isLocked) return;

    // Use today as default slot (current time rounded to nearest 15 min)
    const now = new Date();
    const mins = now.getHours() * 60 + now.getMinutes();
    const slot = Math.floor(mins / SLOT_MINUTES);

    // Default to first team member if task has no assignment
    const resourceId = task.assignedToId || this.teamLanes?.[0]?.id;

    this.closeDrawer();
    this.openSchedulingModal(task, slot, resourceId, null);
  }

  handleUnschedule() {
    // Block unschedule in read-only mode (guest users)
    if (this.readOnlyMode) {
      return;
    }

    const task = this.selectedTask;
    if (!task) return;

    // Project-scoped mode: clear schedule by nulling all 4 fields (assignment remains required).
    if (!this.selectedProjectId) return;

    const assignedToId = task.assignedToId || task.resourceId;
    if (!assignedToId) {
      this.errorMessage = "Assigned To is required.";
      return;
    }

    this.isLoading = true;
    updateActionItemSchedule({
      taskId: task.id,
      assignedToId,
      startDate: null,
      startTime: null,
      endDate: null,
      endTime: null,
      overrideOverlap: false
    })
      .then((res) => {
        if (!res?.success) {
          this.errorMessage = res?.error || "Failed to clear schedule.";
          return null;
        }
        return this.loadDispatchDataForProject(this.selectedProjectId);
      })
      .then(() => {
        this.closeDrawer();
      })
      .catch((err) => {
        console.error("Error clearing schedule", err);
        this.errorMessage = this.formatApexError(err);
      })
      .finally(() => {
        this.isLoading = false;
      });
  }

  // --- Drag & Drop ---
  setDragPayload(e, payload) {
    try {
      if (!e?.dataTransfer || !payload) return;
      e.dataTransfer.setData("task_payload", JSON.stringify(payload));
    } catch (err) {
      console.error("Error setting drag payload", err);
    }
  }

  buildUnscheduledDragPayload(id, source) {
    const ai = this.findUnscheduledActionItemById(id);
    if (!ai) return null;
    if (this.isCompletedStatus(ai.status)) return null;

    return {
      id: ai.id,
      title: ai.title,
      priority: ai.priority,
      status: ai.status,
      assignedToId: ai.assignedToId,
      assignedToName: ai.assignedToName,
      duration: ai.duration,
      source
    };
  }

  buildProjectScheduledDragPayload(id, source) {
    const scheduled = (this.scheduledActionItems || []).find(
      (t) => t.id === id
    );
    if (!scheduled) return null;
    if (this.isCompletedStatus(scheduled.status)) return null;

    return {
      id: scheduled.id,
      title: scheduled.title,
      priority: scheduled.priority,
      status: scheduled.status,
      assignedToId: scheduled.assignedToId,
      assignedToName: scheduled.assignedToName,
      duration: scheduled.duration,
      source
    };
  }

  handleDragStart(e) {
    // Block dragging in read-only mode (guest users)
    if (this.readOnlyMode) {
      e.preventDefault();
      return;
    }
    if (this._isResizing) {
      e.preventDefault();
      return;
    }
    const id = e.currentTarget.dataset.id;
    const source = e.currentTarget.dataset.source;
    if (!id || !source) return;

    // New Dispatch Console: Unscheduled Action Items drag from left panel.
    if (source === "unscheduled") {
      this.setDragPayload(e, this.buildUnscheduledDragPayload(id, source));
    } else if (source === "scheduler" && this.selectedOpportunity) {
      // Project-scoped scheduled task drag (from timeline).
      this.setDragPayload(e, this.buildProjectScheduledDragPayload(id, source));
    }
  }

  handleDragOver(e) {
    e.preventDefault();
    // Provide user feedback for the current drop target.
    if (e.dataTransfer) {
      e.dataTransfer.dropEffect = "move";
    }
  }

  handleDrop(e) {
    e.preventDefault();
    const resId = e.currentTarget.dataset.resourceId;
    const slotIdx = Number(e.currentTarget.dataset.slotIndex);
    if (!Number.isInteger(slotIdx)) return;
    const raw = e.dataTransfer.getData("task_payload");
    if (!raw) return;
    try {
      const payload = JSON.parse(raw);
      // Open scheduling modal instead of immediate drop
      this.dragOverResourceId = null;
      this.dragOverSlotIndex = null;
      this.openSchedulingModal(payload, slotIdx, resId, null);
    } catch (err) {
      console.error("Invalid drag payload", err);
      this.errorMessage = "Unable to schedule: invalid drag payload.";
    }
  }

  handleResizeStart(e) {
    try {
      // Block resizing in read-only mode (guest users)
      if (this.readOnlyMode) {
        return;
      }

      e.preventDefault();
      e.stopPropagation();

      const taskId = e?.currentTarget?.dataset?.id;
      if (!taskId) return;

      const task = this.getScheduledActionItemById(taskId);
      if (!task) return;
      if (this.isCompletedStatus(task.status)) return;

      const rawEnd = this.calculateGlobalSlotIndex(task.endDate, task.endTime);
      if (!Number.isFinite(rawEnd)) return;

      const endSlot = Math.min(this.totalSlots, Math.max(1, rawEnd));

      this._isResizing = true;
      this._resizeTaskId = taskId;
      this._resizeStartClientX = Number(e.clientX) || 0;
      this._resizeStartEndSlot = endSlot;
      this._resizeStartTaskSnapshot = {
        endDate: task.endDate,
        endTime: task.endTime
      };

      // Clear any transient UI messages.
      this.warningMessage = "";
      this.errorMessage = "";
    } catch (err) {
      console.error("Error starting resize", err);
    }
  }

  handleResizeHandleKeyDown(e) {
    // Keep keyboard interaction with the parent block button.
    e.stopPropagation();
  }

  async persistResize(taskId, snapshot) {
    const task = this.getScheduledActionItemById(taskId);
    if (!task) return;

    if (this.isCompletedStatus(task.status)) {
      if (snapshot) this.updateScheduledActionItem(taskId, snapshot);
      return;
    }

    const assignedToId = task.assignedToId || task.resourceId;
    if (!assignedToId) {
      if (snapshot) this.updateScheduledActionItem(taskId, snapshot);
      this.warningMessage = "Assigned To is required.";
      return;
    }

    try {
      this.isLoading = true;

      const result = await updateActionItemSchedule({
        taskId,
        assignedToId,
        startDate: task.startDate,
        startTime: this.normalizeTimeForApex(task.startTime),
        endDate: task.endDate,
        endTime: this.normalizeTimeForApex(task.endTime),
        overrideOverlap: false
      });

      if (result?.success) {
        await this.loadDispatchDataForProject(this.selectedProjectId);
        this.warningMessage = "";
        return;
      }

      if (snapshot) this.updateScheduledActionItem(taskId, snapshot);
      this.warningMessage = result?.error || "Unable to resize this task.";
    } catch (err) {
      console.error("Error saving resize", err);
      if (snapshot) this.updateScheduledActionItem(taskId, snapshot);
      this.errorMessage = this.formatApexError(err);
    } finally {
      this.isLoading = false;
    }
  }

  handleWindowMouseMove(e) {
    if (!this._isResizing) return;
    const taskId = this._resizeTaskId;
    if (!taskId) return;

    const deltaX = (Number(e.clientX) || 0) - (this._resizeStartClientX || 0);
    const deltaSlots = Math.round(deltaX / this.slotWidth); // snap to 15 min
    const newEndSlot = Math.min(
      this.totalSlots,
      Math.max(1, (this._resizeStartEndSlot || 0) + deltaSlots)
    );

    const task = this.getScheduledActionItemById(taskId);
    if (!task) return;

    // Ensure end is after the actual start.
    const rawStart = this.calculateGlobalSlotIndex(
      task.startDate,
      task.startTime
    );
    if (!Number.isFinite(rawStart)) return;
    const minEnd = Math.max(1, rawStart + 1);
    const clampedEndSlot = Math.max(minEnd, newEndSlot);

    const { date: endDate, time: endTime } =
      this.getDateTimeFromSlot(clampedEndSlot);
    if (!endDate || !endTime) return;

    this.updateScheduledActionItem(taskId, { endDate, endTime });
  }

  async handleWindowMouseUp() {
    if (!this._isResizing) return;

    const taskId = this._resizeTaskId;
    const snapshot = this._resizeStartTaskSnapshot;

    this._isResizing = false;
    this._resizeTaskId = null;
    this._resizeStartTaskSnapshot = null;

    if (!taskId) return;

    const task = this.getScheduledActionItemById(taskId);
    if (!task) return;

    // No-op if end didn't change.
    if (
      snapshot &&
      task.endDate === snapshot.endDate &&
      task.endTime === snapshot.endTime
    ) {
      return;
    }

    await this.persistResize(taskId, snapshot);
  }

  handleClearAll() {
    this.selectedProjectId = "";
    this.selectedOpportunity = null;
    this.activeProjectRecord = null;
    this.activeTeam = null;
    this.activeTeamName = null;
    this.teamLanes = [];
    this.targets = [];
    this.actionItems = [];
    this.unscheduledByTarget = [];
    this.sidebarSearchTerm = "";
    this.backlogSearchQuery = "";
    this.isProjectSearchActive = false;
    this.priorityFilter = "All";
    this.durationFilter = "All";
    this.isFilterOpen = false;
    this.selectedTask = null;
  }

  // --- Click-outside handler for dropdowns ---
  handleDocumentClick(e) {
    // Close project dropdown when clicking outside
    if (this.isProjectDropdownOpen) {
      const dropdownContainer = this.template.querySelector(
        ".fs-project-dropdown-container"
      );
      if (dropdownContainer && !dropdownContainer.contains(e.target)) {
        this.isProjectDropdownOpen = false;
      }
    }
  }

  // --- Help modal ---
  handleHelpOpen() {
    this.isHelpModalOpen = true;
  }

  handleHelpClose() {
    this.isHelpModalOpen = false;
  }

  // --- Project Dropdown Search ---
  get selectedProjectLabel() {
    return this.selectedOpportunity
      ? this.selectedOpportunity.Name
      : "Select Project";
  }

  /**
   * Gets filtered project search results for dropdown display
   */
  get filteredProjectResults() {
    return this.projectSearchResults.map((proj) => ({
      ...proj,
      displayLabel: proj.Account?.Name
        ? `${proj.Name} (${proj.Account.Name})`
        : proj.Name
    }));
  }

  handleProjectDropdownOpen() {
    this.isProjectDropdownOpen = true;
    if (!this.projectDropdownSearchTerm) {
      this.loadAllProjects();
    }
  }

  handleProjectDropdownClose() {
    this.isProjectDropdownOpen = false;
  }

  /**
   * Handles project search input in dropdown
   * @param {Event} e - Input event
   */
  async handleProjectDropdownSearch(e) {
    const searchTerm = e.target.value;
    this.projectDropdownSearchTerm = searchTerm;

    if (!searchTerm || searchTerm.length < 2) {
      this.loadAllProjects();
      return;
    }

    try {
      const results = await searchProjects({ searchTerm });
      this.projectSearchResults = results || [];
    } catch (error) {
      console.error("Error searching projects:", error);
      this.projectSearchResults = [];
    }
  }

  /**
   * Loads initial project list for dropdown
   */
  loadAllProjects() {
    this.projectSearchResults = this.allProjects.slice(0, 20);
  }

  async handleProjectSelect(e) {
    const projId = e.currentTarget.dataset.id;
    const selected = this.projectSearchResults.find((p) => p.Id === projId);
    if (!selected) return;

    this.selectedOpportunity = selected;
    this.activeProjectRecord = selected;
    this.isProjectDropdownOpen = false;
    this.projectDropdownSearchTerm = "";

    // Single active Project drives everything.
    this.selectedProjectId = projId;

    await this.loadDispatchDataForProject(projId);
  }

  // Helper to format time for display
  formatTimeDisplay(timeValue) {
    if (!timeValue) return "Not set";
    // timeValue is in format "HH:mm:ss.sssZ" or similar
    const parts = timeValue.split(":");
    if (parts.length >= 2) {
      const hour = Number(parts[0]);
      const minute = parts[1];
      const ampm = hour >= 12 ? "PM" : "AM";
      const hour12 = hour % 12 || 12;
      return `${hour12}:${minute} ${ampm}`;
    }
    return timeValue;
  }

  // Helper to calculate duration display
  formatDuration(minutes) {
    if (!minutes) return "0m";
    const hours = Math.floor(minutes / 60);
    const mins = Math.round(minutes % 60);
    if (hours > 0) {
      return mins > 0 ? `${hours}h ${mins}m` : `${hours}h`;
    }
    return `${mins}m`;
  }

  // Open scheduling modal when task is dropped
  openSchedulingModal(task, slotIndex, resourceId, existingTaskId) {
    // Completed/closed action items are locked: keep visible, but do not open modal.
    if (this.isCompletedStatus(task?.status)) return;

    // Calculate date/time from slot index
    const { date, time } = this.getDateTimeFromSlot(slotIndex);

    this.schedulingTask = {
      ...task,
      targetSlotIndex: slotIndex,
      targetResourceId: resourceId,
      targetTaskId: existingTaskId
    };

    // Assignment is required. Default to drop lane, then task's current assignment.
    this.scheduleAssignedToId = resourceId || task?.assignedToId || null;

    this.scheduleStartDate = date;
    this.scheduleStartTime = time;

    // Default end time based on duration
    const duration = task.duration || task.calculatedDuration || 60;
    const endDateTime = this.calculateEndDateTime(date, time, duration);
    this.scheduleEndDate = endDateTime.date;
    this.scheduleEndTime = endDateTime.time;
    this.estimatedDuration = duration;

    this.isSchedulingModalOpen = true;
  }

  closeSchedulingModal() {
    this.isSchedulingModalOpen = false;
    this.schedulingTask = null;
    this.scheduleAssignedToId = null;
    this.scheduleStartDate = null;
    this.scheduleStartTime = null;
    this.scheduleEndDate = null;
    this.scheduleEndTime = null;
    this.estimatedDuration = null;
  }

  handlePrevDay() {
    const next = addDaysToDateStr(this.activeDateStr, -1);
    if (next) this.activeDateStr = next;
  }

  handleNextDay() {
    const next = addDaysToDateStr(this.activeDateStr, 1);
    if (next) this.activeDateStr = next;
  }

  handleToday() {
    this.activeDateStr = todayInTimeZoneDateStr(TIME_ZONE);
  }

  async handleRefresh() {
    try {
      if (this.selectedProjectId) {
        await this.loadDispatchDataForProject(this.selectedProjectId);
        return;
      }
      await this.loadInitialData();
    } catch (err) {
      console.error("Error refreshing", err);
      this.errorMessage = this.formatApexError(err);
    }
  }

  handleZoomOut() {
    const idx = Math.max(0, Number(this.zoomIndex) - 1);
    this.zoomIndex = idx;
  }

  handleZoomIn() {
    const idx = Math.min(ZOOM_LEVELS.length - 1, Number(this.zoomIndex) + 1);
    this.zoomIndex = idx;
  }

  handleTimelineRangeToggle(e) {
    const mode = e?.currentTarget?.dataset?.mode;
    if (!mode) return;
    this.timelineRange = mode;
  }

  get timelineRangeIsDay() {
    return this.timelineRange === "day";
  }

  get timelineRangeIsMultiDay() {
    return this.timelineRange === "multiday";
  }

  get timelineRangeIsWeek() {
    return this.timelineRange === "week";
  }

  get dayRangeButtonClass() {
    return `dc-seg ${this.timelineRangeIsDay ? "dc-seg-active" : ""}`;
  }

  get multiDayRangeButtonClass() {
    return `dc-seg ${this.timelineRangeIsMultiDay ? "dc-seg-active" : ""}`;
  }

  get weekRangeButtonClass() {
    return `dc-seg ${this.timelineRangeIsWeek ? "dc-seg-active" : ""}`;
  }

  handleGlobalSearchInput(e) {
    this.globalSearchQuery = e?.target?.value || "";
  }

  handleScheduleAssignedToChange(e) {
    this.scheduleAssignedToId = e.detail.value;
  }

  handleScheduleStartDateChange(e) {
    this.scheduleStartDate = e.target.value;
    this.recalculateScheduleDuration();
  }

  handleScheduleStartTimeChange(e) {
    this.scheduleStartTime = e.target.value;
    this.recalculateScheduleDuration();
  }

  handleScheduleEndDateChange(e) {
    this.scheduleEndDate = e.target.value;
    this.recalculateScheduleDuration();
  }

  handleScheduleEndTimeChange(e) {
    this.scheduleEndTime = e.target.value;
    this.recalculateScheduleDuration();
  }

  recalculateScheduleDuration() {
    if (
      !this.scheduleStartDate ||
      !this.scheduleStartTime ||
      !this.scheduleEndDate ||
      !this.scheduleEndTime
    ) {
      return;
    }

    try {
      // Parse dates manually to avoid timezone shifts (YYYY-MM-DD)
      const startDateParts = this.scheduleStartDate.split("-");
      const endDateParts = this.scheduleEndDate.split("-");
      const startTimeParts = this.scheduleStartTime.split(":");
      const endTimeParts = this.scheduleEndTime.split(":");

      const startDateTime = new Date(
        Number(startDateParts[0]),
        Number(startDateParts[1]) - 1,
        Number(startDateParts[2]),
        Number(startTimeParts[0]),
        Number(startTimeParts[1])
      );

      const endDateTime = new Date(
        Number(endDateParts[0]),
        Number(endDateParts[1]) - 1,
        Number(endDateParts[2]),
        Number(endTimeParts[0]),
        Number(endTimeParts[1])
      );

      const durationMs = endDateTime.getTime() - startDateTime.getTime();
      this.estimatedDuration = Math.round(durationMs / (1000 * 60));
    } catch (error) {
      console.error("Error calculating duration:", error);
    }
  }

  get scheduleValidationError() {
    if (!this.isSchedulingModalOpen) return "";

    if (!this.scheduleAssignedToId) {
      return "Assigned To is required.";
    }

    if (
      !this.scheduleStartDate ||
      !this.scheduleStartTime ||
      !this.scheduleEndDate ||
      !this.scheduleEndTime
    ) {
      return "";
    }
    const mins = Number(this.estimatedDuration);
    if (!Number.isFinite(mins)) return "";
    if (mins <= 0) return "End date/time must be after start date/time.";
    return "";
  }

  get isConfirmScheduleDisabled() {
    return Boolean(
      this.readOnlyMode || this.isLoading || !this.schedulingTask || this.scheduleValidationError
    );
  }

  normalizeTimeForApex(timeValue) {
    if (!timeValue) return null;
    const t = String(timeValue).trim();

    // Required: Send "HH:mm:ss" string to Apex (not Time object or ISO string).
    // The Apex controller will parse this string.

    // HH:mm -> HH:mm:00
    if (/^\d{1,2}:\d{2}$/.test(t)) return `${t}:00`;

    // HH:mm:ss -> HH:mm:ss
    if (/^\d{1,2}:\d{2}:\d{2}$/.test(t)) return t;

    // HH:mm:ss.SSSZ or HH:mm:ss.SSS -> strip milliseconds
    if (t.includes(".")) {
      return t.substring(0, t.indexOf("."));
    }
    // HH:mm:ssZ -> strip Z
    if (t.endsWith("Z")) {
      return t.substring(0, t.length - 1);
    }

    return t;
  }

  /**
   * Handles schedule confirmation from modal
   */
  async handleConfirmSchedule() {
    if (!this.schedulingTask) return;

    if (this.scheduleValidationError) {
      this.errorMessage = this.scheduleValidationError;
      return;
    }

    try {
      this.isLoading = true;

      const payload = {
        taskId: this.schedulingTask.id,
        assignedToId: this.scheduleAssignedToId,
        startDate: this.scheduleStartDate,
        startTime: this.normalizeTimeForApex(this.scheduleStartTime),
        endDate: this.scheduleEndDate,
        endTime: this.normalizeTimeForApex(this.scheduleEndTime),
        overrideOverlap: false
      };

      const result = await updateActionItemSchedule(payload);

      if (result.success) {
        await this.loadDispatchDataForProject(this.selectedProjectId);
        this.closeSchedulingModal();
        this.isLoading = false;
        return;
      }

      this.errorMessage = result.error;
      this.isLoading = false;
    } catch (error) {
      console.error("Error scheduling task:", error);
      this.errorMessage = this.formatApexError(error);
      this.isLoading = false;
    }
  }

  applyScheduledTaskToTeamState() {
    // Legacy mode removed; project-scoped refresh from Apex is the source of truth.
  }

  getDateTimeFromSlot(slotIndex) {
    const start = this.activeDateStr;
    if (!start || !Number.isInteger(slotIndex)) {
      return { date: null, time: null };
    }

    const dayOffset = Math.floor(slotIndex / SLOTS_PER_DAY);
    const slotInDay =
      ((slotIndex % SLOTS_PER_DAY) + SLOTS_PER_DAY) % SLOTS_PER_DAY;
    const minutesFromMidnight = slotInDay * SLOT_MINUTES;
    const hour = Math.floor(minutesFromMidnight / 60);
    const minute = minutesFromMidnight % 60;

    const date = addDaysToDateStr(start, dayOffset);
    const time = `${pad2(hour)}:${pad2(minute)}`;
    return { date, time };
  }

  calculateEndDateTime(dateStr, timeStr, durationMinutes) {
    try {
      const tParts = String(timeStr || "").split(":");
      const hour = Number(tParts[0]);
      const minute = Number(tParts[1]);
      if (!Number.isFinite(hour) || !Number.isFinite(minute)) {
        return { date: dateStr, time: timeStr };
      }

      const startTotal = hour * 60 + minute;
      const endTotal = startTotal + Number(durationMinutes || 0);
      const daysToAdd = Math.floor(endTotal / (24 * 60));
      const endMins = ((endTotal % (24 * 60)) + 24 * 60) % (24 * 60);
      const endHour = Math.floor(endMins / 60);
      const endMinute = endMins % 60;

      const endDate = addDaysToDateStr(dateStr, daysToAdd) || dateStr;
      const endTime = `${pad2(endHour)}:${pad2(endMinute)}`;
      return { date: endDate, time: endTime };
    } catch (error) {
      console.error("Error calculating end time:", error);
      return { date: dateStr, time: timeStr };
    }
  }

  calculateGlobalSlotIndex(dateStr, timeStr) {
    const start = this.activeDateStr;
    if (!start || !dateStr || !timeStr) return null;

    const dayOffset = diffDays(dateStr, start);
    if (!Number.isFinite(dayOffset)) return null;

    const tParts = String(timeStr).split(":");
    if (tParts.length < 2) return null;
    const hour = Number(tParts[0]);
    const minute = Number(tParts[1]);
    if (!Number.isFinite(hour) || !Number.isFinite(minute)) return null;

    const slotInDay = Math.floor((hour * 60 + minute) / SLOT_MINUTES);
    return dayOffset * SLOTS_PER_DAY + slotInDay;
  }

  // Convert a date/time to a slot index within the current visible range.
  calculateSlotIndexFromDateTime(dateStr, timeStr) {
    const idx = this.calculateGlobalSlotIndex(dateStr, timeStr);
    if (!Number.isFinite(idx)) return null;
    if (idx < 0 || idx >= this.totalSlots) return null;
    return idx;
  }

  // Enrich task with positioning for timeline display
  enrichTaskForTimeline(task) {
    if (!task.isScheduled) {
      return { ...task, isScheduled: false };
    }

    const isLocked = this.isCompletedStatus(task?.status);

    const rawStart = this.calculateGlobalSlotIndex(
      task.startDate,
      task.startTime
    );
    const rawEnd = this.calculateGlobalSlotIndex(task.endDate, task.endTime);
    if (!Number.isFinite(rawStart) || !Number.isFinite(rawEnd)) {
      // Task is scheduled but outside visible range - keep it as scheduled
      // but mark it as not renderable on current timeline view
      return {
        ...task,
        isScheduled: true,
        isVisible: false,
        blockStyle: "display: none;"
      };
    }

    const startIdx = Math.max(0, rawStart);
    const endIdx = Math.min(this.totalSlots, rawEnd);
    const durationSlots = Math.max(1, endIdx - startIdx);

    const width = durationSlots * this.slotWidth;
    const left = startIdx * this.slotWidth;

    const start = this.getSlotTime(startIdx);
    const end = this.getSlotTime(endIdx);
    const priority = task.priority || "Medium";

    // Calculate actual duration display
    const duration = task.calculatedDuration || task.duration || 60;

    const spansMultipleDays = String(task.startDate) !== String(task.endDate);
    const isContinued = rawStart < 0;
    const isTruncatedEnd = rawEnd > this.totalSlots;

    // Repeat the label at each day boundary inside a single continuous bar.
    const labelSegments = [];
    try {
      const title = String(task.title || "");
      const firstDay = Math.floor(startIdx / SLOTS_PER_DAY);
      const lastDay = Math.floor(
        (Math.max(startIdx + 1, endIdx) - 1) / SLOTS_PER_DAY
      );
      for (let d = firstDay; d <= lastDay; d++) {
        const segStart = Math.max(startIdx, d * SLOTS_PER_DAY);
        const leftPx = Math.max(0, (segStart - startIdx) * this.slotWidth);
        labelSegments.push({
          key: `${task.id}-${d}`,
          text: title,
          style: `left: ${leftPx + 8}px;`
        });
      }
    } catch {
      // no-op
    }

    return {
      ...task,
      startSlotIndex: startIdx,
      endSlotIndex: endIdx,
      timeRange: `${start} - ${end}`,
      blockClass: `scheduled-block ${priority.toLowerCase()} ${isContinued || isTruncatedEnd ? "continued-task" : ""} ${task.showOverlapIndicator ? "has-overlap" : ""} ${isLocked ? "is-locked" : ""}`,
      blockStyle: `left: ${left + 6}px; width: ${Math.max(0, width - 12)}px;`,
      priorityClass: `badge-priority ${priority.toLowerCase()}`,
      durationDisplay: this.formatDuration(duration),
      isScheduled: true,
      isLocked,
      isDraggable: !isLocked && !this.readOnlyMode,
      isMultiDay: spansMultipleDays,
      isContinued,
      showOverlapIndicator: Boolean(task.showOverlapIndicator),
      labelSegments,
      ariaLabel: `${task.title || "Task"}. ${start} to ${end}. ${isContinued ? "Continued from previous range." : ""}`
    };
  }
}