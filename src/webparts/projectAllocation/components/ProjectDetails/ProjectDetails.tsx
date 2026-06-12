import * as React from 'react';
import styles from './projectdetails.module.scss';
import type { ProjectRow } from '../Dashboard/dashboard';
import { getSP } from '../../../../service/initservice';
import { useUser } from '../UserContext/UserContext';

// ─── Types ────────────────────────────────────────────────────────────────────

type DetailTab   = 'tasks' | 'logs';
type TaskStatus  = 'open' | 'in_progress' | 'review' | 'completed' | 'blocked';
type TaskPriority = 'low' | 'medium' | 'high' | 'critical';

interface TaskRecord {
  id: string;
  name: string;
  associated: string;
  tags: string[];
  owner: string;
  ownerInitials: string;
  priority: TaskPriority;
  status: TaskStatus;
  startDate: string;   // SP StartDate
  dueDate: string;     // SP EndDate
  progress: number;
  workHours: string;
  timeLog: string;
  difference: string;
  billingType: string;
  effort: string;
  lane: string;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

const VALID_STATUSES: TaskStatus[] = ['open', 'in_progress', 'review', 'completed', 'blocked'];
const VALID_PRIORITIES: TaskPriority[] = ['low', 'medium', 'high', 'critical'];

const normaliseStatus = (raw: string): TaskStatus => {
  const s = (raw || '').toLowerCase().replace(/\s+/g, '_');
  if (s === 'not_started' || s === 'not started') return 'open';
  if (s === 'in_progress' || s === 'in progress') return 'in_progress';
  if (s === 'review' || s === 'in_review')         return 'review';
  if (s === 'done' || s === 'completed')           return 'completed';
  if (s === 'blocked' || s === 'on_hold')          return 'blocked';
  return VALID_STATUSES.indexOf(s as TaskStatus) >= 0 ? (s as TaskStatus) : 'open';
};

const normalisePriority = (raw: string): TaskPriority => {
  const p = (raw || '').toLowerCase();
  return VALID_PRIORITIES.indexOf(p as TaskPriority) >= 0 ? (p as TaskPriority) : 'medium';
};

const getInitials = (name: string): string => {
  if (!name) return '??';
  const parts = name.trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
};

const formatDate = (dateStr: string): string => {
  if (!dateStr) return '—';
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('en-US', { month: 'short', day: '2-digit', year: 'numeric' });
  } catch {
    return dateStr;
  }
};

const getFieldValue = (item: any, fieldName: string): any => {
  if (!item) return undefined;
  const lower = fieldName.toLowerCase();
  for (const key of Object.keys(item)) {
    if (key.toLowerCase() === lower) {
      const val = item[key];
      if (val && typeof val === 'object') {
        if (val.LookupValue !== undefined) return val.LookupValue;
        if (val.Title !== undefined) return val.Title;
        if (val.Label !== undefined) return val.Label;
      }
      return val;
    }
  }
  return undefined;
};

const normalizeProjectCode = (code: unknown): string =>
  (code ?? '').toString().trim().toLowerCase();

const escapeODataString = (value: string): string =>
  value.replace(/'/g, "''");

const projectCodesMatch = (selectedCode: string, taskCode: unknown): boolean => {
  const a = normalizeProjectCode(selectedCode);
  const b = normalizeProjectCode(taskCode);
  if (!a || !b) return false;
  if (a === b) return true;

  if (a.indexOf('prj-') === 0 && b === a.replace('prj-', '')) return true;
  if (b.indexOf('prj-') === 0 && a === b.replace('prj-', '')) return true;
  return false;
};

const progressFromStatus = (status: TaskStatus): number => {
  if (status === 'completed') return 100;
  if (status === 'review') return 75;
  if (status === 'in_progress') return 50;
  if (status === 'blocked') return 10;
  return 0;
};

const mapSpItemToTask = (item: any, projectCode: string, idx: number): TaskRecord => {
  const owner = String(getFieldValue(item, 'Assignee') || 'Unassigned');
  const taskName = String(getFieldValue(item, 'Task') || `Task ${idx + 1}`);
  const sno = String(getFieldValue(item, 'Sno') || idx + 1);
  const status = normaliseStatus(String(getFieldValue(item, 'Status') || ''));
  const lane = String(getFieldValue(item, 'lane') || '');
  const effort = getFieldValue(item, 'Effort');
  const effortText = effort !== undefined && effort !== null && effort !== '' ? String(effort) : '0';
  const startD = String(getFieldValue(item, 'StartDate') || '');
  const endD = String(getFieldValue(item, 'EndDate') || '');

  return {
    id: `${projectCode}-T${sno.length < 2 ? `0${sno}` : sno}`,
    name: taskName,
    associated: lane ? `Lane: ${lane}` : 'Not Associated',
    tags: lane ? [lane] : [],
    owner,
    ownerInitials: getInitials(owner),
    priority: normalisePriority(String(getFieldValue(item, 'Priority') || '')),
    status,
    startDate: formatDate(startD),
    dueDate: formatDate(endD),
    progress: progressFromStatus(status),
    workHours: `${effortText}h`,
    timeLog: '00:00',
    difference: effortText !== '0' ? `- ${effortText}h` : '00:00',
    billingType: 'None',
    effort: effortText,
    lane: lane || 'development',
  };
};

const TABS: { id: DetailTab; label: string; icon: 'tasks' | 'clock' }[] = [
  { id: 'tasks', label: 'Tasks',    icon: 'tasks' },
  { id: 'logs',  label: 'Log Hours', icon: 'clock' }
];

const STATUS_LABEL: Record<TaskStatus, string> = {
  open:        'Open',
  in_progress: 'In Progress',
  review:      'Review',
  completed:   'Completed',
  blocked:     'Blocked'
};

const GROUP_OPTIONS = ['Task List (General)', 'By Status', 'By Owner', 'By Priority'];


export interface IProjectDetailsProps {
  project: ProjectRow;
  onBack: () => void;
}


const TabIcon: React.FC<{ type: 'tasks' | 'clock'; className?: string }> = ({ type, className }) => {
  if (type === 'clock') {
    return (
      <svg viewBox="0 0 24 24" className={className} aria-hidden="true">
        <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="2" fill="none" />
        <path d="M12 7v5l3 2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" fill="none" />
      </svg>
    );
  }
  return (
    <svg viewBox="0 0 24 24" className={className} aria-hidden="true">
      <rect x="4" y="4" width="16" height="16" rx="3" stroke="currentColor" strokeWidth="2" fill="none" />
      <path d="M8 10h8M8 14h5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
    </svg>
  );
};


const ProjectDetails: React.FC<IProjectDetailsProps> = ({ project, onBack }) => {
  const css = styles as typeof styles & Record<string, string>;
  const { displayName, email, initials, greeting } = useUser();

  const [tab,         setTab]         = React.useState<DetailTab>('tasks');
  const [groupBy,     setGroupBy]     = React.useState(GROUP_OPTIONS[0]);
  const [selectedIds, setSelectedIds] = React.useState<Set<string>>(new Set());

  const [tasks,        setTasks]    = React.useState<TaskRecord[]>([]);
  const [loading,      setLoading]  = React.useState(true);


  const fetchTasks = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setTasks([]);
      const sp = getSP();
      const projectCode = (project.id || '').trim();

      const taskFields = [
        'Id', 'Task', 'Sno', 'ProjectCode', 'Assignee', 'lane',
        'Effort', 'Status', 'StartDate', 'EndDate'
      ];

      let spItems: any[] = [];
      if (projectCode) {
        try {
          spItems = await sp.web.lists
            .getByTitle('TaskDetails')
            .items
            .select(...taskFields)
            .filter(`ProjectCode eq '${escapeODataString(projectCode)}'`)
            .top(5000)();
        } catch (filterErr) {
          console.warn('[ProjectDetails] OData filter failed, falling back to client filter:', filterErr);
        }
      }
      if (spItems.length === 0 && projectCode) {
        const allItems: any[] = await sp.web.lists
          .getByTitle('TaskDetails')
          .items
          .select(...taskFields)
          .top(5000)();

        spItems = allItems.filter((item) =>
          projectCodesMatch(projectCode, getFieldValue(item, 'ProjectCode'))
        );
      }

      const matchedItems = spItems;

      matchedItems.sort((a, b) => {
        const snoA = parseInt(getFieldValue(a, 'Sno') || '0', 10);
        const snoB = parseInt(getFieldValue(b, 'Sno') || '0', 10);
        return snoA - snoB;
      });

      setTasks(matchedItems.map((item, idx) => mapSpItemToTask(item, projectCode, idx)));

    } catch (error) {
      console.error('[ProjectDetails] Error fetching tasks from SharePoint:', error);
      setTasks([]);
    } finally {
      setLoading(false);
    }
  }, [project.id]);

  React.useEffect(() => {
    fetchTasks().catch(console.error);
  }, [fetchTasks]);

  
  const toggleRow = (id: string): void => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  const toggleAll = (): void => {
    if (selectedIds.size === tasks.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(tasks.map(t => t.id)));
    }
  };

  
  const displayedTasks = React.useMemo(() => {
    if (groupBy === 'By Status') {
      const order: TaskStatus[] = ['in_progress', 'review', 'open', 'blocked', 'completed'];
      return [...tasks].sort((a, b) => order.indexOf(a.status) - order.indexOf(b.status));
    }
    if (groupBy === 'By Owner') {
      return [...tasks].sort((a, b) => a.owner.localeCompare(b.owner));
    }
    if (groupBy === 'By Priority') {
      const order: TaskPriority[] = ['critical', 'high', 'medium', 'low'];
      return [...tasks].sort((a, b) => order.indexOf(a.priority) - order.indexOf(b.priority));
    }
    return tasks;
  }, [tasks, groupBy]);

  
  const renderTasksTable = (): React.ReactNode => {
    if (loading) {
      return (
        <div className={css.loadingWrap} aria-live="polite">
          <div className={css.spinnerRing} />
          <span>Loading tasks for project <strong>{project.id}</strong>…</span>
        </div>
      );
    }

    if (displayedTasks.length === 0) {
      return (
        <div className={css.emptyState}>
          <svg viewBox="0 0 24 24" className={css.emptyIcon} fill="none" aria-hidden="true">
            <rect x="4" y="4" width="16" height="16" rx="3" stroke="currentColor" strokeWidth="1.8" />
            <path d="M8 10h8M8 14h5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
          </svg>
          <p className={css.emptyTitle}>No tasks found</p>
          <p className={css.emptyDesc}>No tasks in <strong>TaskDetails</strong> match project code <code className={css.codeTag}>{project.id}</code></p>
        </div>
      );
    }

    return (
      <div className={css.tableCard}>
        <div className={css.tableCardGlow} aria-hidden="true" />
        <div className={css.gridWrap}>
          <table className={css.dataGrid}>
            <thead>
              <tr>
                <th className={css.colCheck}>
                  <label className={css.checkLabel}>
                    <input
                      type="checkbox"
                      checked={selectedIds.size === displayedTasks.length && displayedTasks.length > 0}
                      onChange={toggleAll}
                      aria-label="Select all tasks"
                    />
                    <span className={css.checkBox} />
                  </label>
                </th>
                <th className={css.colTaskName}>Task</th>
                <th className={css.colEffort}>Effort (h)</th>
                <th className={css.colOwner}>Assignee</th>
                <th className={css.colDate}>Start Date</th>
                <th className={css.colDate}>End Date</th>
                <th className={css.colStatus}>Status</th>
              </tr>
            </thead>
            <tbody>
              {displayedTasks.map((task) => {
                const selected = selectedIds.has(task.id);
                return (
                  <tr key={task.id} className={selected ? css.rowSelected : undefined}>
                   
                   
                    <td className={css.colCheck}>
                      <label className={css.checkLabel}>
                        <input
                          type="checkbox"
                          checked={selected}
                          onChange={() => toggleRow(task.id)}
                          aria-label={`Select ${task.name}`}
                        />
                        <span className={css.checkBox} />
                      </label>
                    </td>

                   
                   
                    <td className={css.colTaskName}>
                      <div className={css.taskNameCell}>
                        <span
                          className={`${css.statusIndicator} ${css[`status_${task.status}`]}`}
                          title={STATUS_LABEL[task.status]}
                        />
                        <div>
                          <span className={css.taskTitle}>{task.name}</span>
                          {task.lane && (
                            <span className={css.laneTag}>{task.lane}</span>
                          )}
                        </div>
                      </div>
                    </td>

                    {/* Effort */}
                    <td className={css.colEffort}>
                      <span className={css.effortBadge}>{task.effort}h</span>
                    </td>

                    {/* Assignee */}
                    <td className={css.colOwner}>
                      <div className={css.assigneeCell}>
                        <span className={css.ownerAvatar}>{task.ownerInitials}</span>
                        <span className={css.ownerName}>{task.owner}</span>
                      </div>
                    </td>

                    {/* Start Date */}
                    <td className={css.colDate}>
                      <span className={css.dateVal}>{task.startDate || '—'}</span>
                    </td>

                    {/* End Date */}
                    <td className={css.colDate}>
                      <span className={css.dateVal}>{task.dueDate}</span>
                    </td>

                    {/* Status */}
                    <td className={css.colStatus}>
                      <span className={`${css.statusChip} ${css[`status_${task.status}`]}`}>
                        <span className={css.statusDot} />
                        {STATUS_LABEL[task.status]}
                      </span>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    );
  };


  const totalEffortHrs = displayedTasks.reduce((s, t) => s + (parseInt(t.effort, 10) || 0), 0);

  const renderLogHours = (): React.ReactNode => {
    if (loading) {
      return (
        <div className={css.loadingWrap} aria-live="polite">
          <div className={css.spinnerRing} />
          <span>Loading work logs…</span>
        </div>
      );
    }

    return (
      <div className={css.tableCard}>
        <div className={css.tableCardGlow} aria-hidden="true" />
        <div className={css.logsBanner}>
          <span>Project Work Logs &amp; Effort</span>
          <span className={css.totalWarn}>Total Effort: {totalEffortHrs}h</span>
        </div>
        <div className={css.gridWrap}>
          <table className={css.dataGrid}>
            <thead>
              <tr>
                <th className={css.colName}>Task</th>
                <th className={css.colHours}>Effort (h)</th>
                <th className={css.colHours}>Timelog Total</th>
                <th className={css.colDiff}>Difference</th>
                <th className={css.colBilling}>Billing Type</th>
              </tr>
            </thead>
            <tbody>
              {displayedTasks.map((task) => (
                <tr key={task.id}>
                  <td className={css.colName}>
                    <div className={css.taskCellCompact}>
                      <span className={css.taskTitle}>{task.name}</span>
                      <span className={css.taskIdMeta}>{task.id}</span>
                    </div>
                  </td>
                  <td className={css.colHours}><span className={css.cellMuted}>{task.workHours}</span></td>
                  <td className={css.colHours}><span className={css.boldVal}>{task.timeLog}</span></td>
                  <td className={css.colDiff}>
                    <span className={task.difference.startsWith('-') ? css.diffBad : css.diffNeutral}>
                      {task.difference}
                    </span>
                  </td>
                  <td className={css.colBilling}><span className={css.billingText}>{task.billingType}</span></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  
  return (
    <div className={css.page}>
      <div className={css.pageMesh}  aria-hidden="true" />
      <div className={css.pageOrb1} aria-hidden="true" />
      <div className={css.pageOrb2} aria-hidden="true" />

      {/* ── Sticky header ── */}
      <header className={css.stickyHeader}>
        <div className={css.headerLeft}>
          <button type="button" className={css.backBtn} onClick={onBack} aria-label="Back to projects">
            <svg viewBox="0 0 24 24" className={css.backIcon} aria-hidden="true">
              <path d="M15 19l-7-7 7-7" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" fill="none" />
            </svg>
          </button>
          <div className={css.headerTitles}>
            <h1 className={css.appTitle}>Project Task Manager</h1>
            <p className={css.projectMeta}>
              <span className={css.projectIdPill}>{project.id}</span>
              <span className={css.metaSep}>·</span>
              {project.name}
              {loading && <span className={css.loadingBadge}>⟳ Loading tasks…</span>}
            </p>
          </div>
        </div>

        <div className={css.headerRight}>
          <div className={css.searchWrap}>
            <svg viewBox="0 0 24 24" className={css.searchSvg} aria-hidden="true">
              <circle cx="11" cy="11" r="7" stroke="currentColor" strokeWidth="2.2" fill="none" />
              <path d="M20 20l-4.3-4.3" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" fill="none" />
            </svg>
            <input placeholder="Search core intelligence..." aria-label="Search" />
          </div>
          <button type="button" className={css.userMenu} aria-label="User menu">
            <div className={css.userMenuText}>
              <strong>{greeting}, {displayName}</strong>
              {email ? <small>{email}</small> : null}
            </div>
            <span className={css.avatar}>{initials}</span>
          </button>
        </div>
      </header>

      {/* ── Body ── */}
      <div className={css.bodyLayout}>
        <main className={css.mainColumn}>

          {/* Tabs */}
          <nav className={css.viewTabs} aria-label="Project views">
            {TABS.map((t) => (
              <button
                key={t.id}
                type="button"
                className={tab === t.id ? css.tabActive : css.tabBtn}
                onClick={() => setTab(t.id)}
              >
                <TabIcon type={t.icon} className={css.tabIcon} />
                {t.label}
              </button>
            ))}
          </nav>

          {/* Toolbar */}
          <div className={css.contentToolbar}>
            <label className={css.groupByWrap}>
              <span className={css.groupByLabel}>Group By:</span>
              <select value={groupBy} onChange={(e) => setGroupBy(e.target.value)} aria-label="Group by">
                {GROUP_OPTIONS.map((opt) => (
                  <option key={opt} value={opt}>{opt}</option>
                ))}
              </select>
            </label>
            <div className={css.toolbarRight}>
              <span className={css.taskCount}>
                Showing <strong>{displayedTasks.length}</strong> task{displayedTasks.length !== 1 ? 's' : ''}
              </span>
              {tab === 'tasks' && (
                <button type="button" className={css.newTaskBtn}>+ New Task</button>
              )}
            </div>
          </div>

          {tab === 'tasks' && renderTasksTable()}
          {tab === 'logs'  && renderLogHours()}
        </main>
      </div>
    </div>
  );
};

export default ProjectDetails;
