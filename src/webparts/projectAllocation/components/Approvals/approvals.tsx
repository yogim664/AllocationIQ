import * as React from 'react';
import * as XLSX from 'xlsx';
import styles from './approvals.module.scss';
import NavBar from '../NavBar/navbar';
import { useUser } from '../UserContext/UserContext';
import { getSP } from '../../../../service/initservice';

interface StaffMember {
  initials: string;
  name: string;
  role: string;
  department: string;
  match: number;
  availability: number;
  success: number;
}

interface PlanItem {
  slNo: number;
  feature: string;
  task: string;
  startDate: string;
  endDate: string;
  efforts: number;
  lane: 'development' | 'demo' | 'testing';
  assignee: string;
  status: 'Not Started' | 'In Progress' | 'Done';
}

export interface ApprovalProject {
  spItemId: number;
  id: string;
  name: string;
  subtitle: string;
  client: string;
  budget: string;
  startDate: string;
  endDate: string;
  priority: 'Low' | 'Medium' | 'High' | 'Critical';
  skills: string[];
  team: string[];
  teamExtra: number;
  status: 'Pending' | 'Approved' | 'Rejected';
  staff: StaffMember[];
  planItems: PlanItem[];
 
  modified?: string;
}


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

const getManagerApprovedValue = (item: any): boolean | null => {
  const raw =
    item.isManagerApproved ??
    item.IsManagerApproved ??
    getFieldValue(item, 'isManagerApproved') ??
    getFieldValue(item, 'IsManagerApproved');

  if (raw === true || raw === 1 || raw === '1' || raw === 'true' || raw === 'Yes') return true;
  if (raw === false || raw === 0 || raw === '0' || raw === 'false' || raw === 'No') return false;
  return null;
};


const isManagerApprovedFalse = (item: any): boolean =>
  getManagerApprovedValue(item) === false;

const projectCodesMatch = (a: string, b: string): boolean =>
  (a || '').trim().toLowerCase() === (b || '').trim().toLowerCase();

const escapeODataString = (value: string): string =>
  value.replace(/'/g, "''");

const PARTICIPATION_SELECT = [
  'Id', 'field_1', 'field_2', 'field_3', 'field_4', 'field_5', 'field_6', 'field_7'
];

const readTextField = (item: any, ...names: string[]): string => {
  for (const name of names) {
    const val = getFieldValue(item, name) ?? item[name];
    if (val == null || val === '') continue;
    if (typeof val === 'object') {
      const text = val.LookupValue ?? val.Title ?? val.Label ?? val.value;
      if (text != null && text !== '') return String(text).trim();
    } else {
      return String(val).trim();
    }
  }
  return '';
};

const getParticipationProjectCode = (item: any): string =>
  readTextField(item, 'field_3', 'ProjectCode', 'projectcode');

const getParticipationName = (item: any): string =>
  readTextField(item, 'field_2', 'FullName', 'Title');

const formatDate = (dateStr: string): string => {
  if (!dateStr) return '';
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('en-US', { month: 'short', day: '2-digit', year: 'numeric' });
  } catch {
    return dateStr;
  }
};

const getInitials = (fullName: string): string => {
  if (!fullName) return '??';
  const parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
};

const parseModifiedTime = (value?: string): number => {
  if (!value) return 0;
  const time = new Date(value).getTime();
  return isNaN(time) ? 0 : time;
};


const sortProjectsByModifiedDesc = (items: ApprovalProject[]): ApprovalProject[] =>
  [...items].sort((a, b) => {
    const diff = parseModifiedTime(b.modified) - parseModifiedTime(a.modified);
    if (diff !== 0) return diff;
    return b.id.localeCompare(a.id, undefined, { numeric: true, sensitivity: 'base' });
  });

const normaliseLane = (raw: string): PlanItem['lane'] => {
  const lane = (raw || '').toLowerCase();
  if (lane === 'testing') return 'testing';
  if (lane === 'demo') return 'demo';
  return 'development';
};

const normalisePlanStatus = (raw: string): PlanItem['status'] => {
  const s = (raw || '').toLowerCase();
  if (s === 'in progress' || s === 'in_progress') return 'In Progress';
  if (s === 'done' || s === 'completed') return 'Done';
  return 'Not Started';
};

const mapParticipationToStaff = (item: any): StaffMember => {
  const name = getParticipationName(item);
  const availability = parseInt(
    readTextField(item, 'field_5', 'ContributionPercent') || '75',
    10
  ) || 75;
  const role = (readTextField(item, 'field_4', 'RoleInProject') || 'TEAM MEMBER').toUpperCase();
  return {
    initials: getInitials(name),
    name,
    role,
    department: 'Project Team',
    match: Math.min(99, availability + 10),
    availability,
    success: 90
  };
};

const mapTaskToPlanItem = (t: any, idx: number): PlanItem => ({
  slNo: parseInt(String(getFieldValue(t, 'Sno') || idx + 1), 10) || idx + 1,
  feature: (getFieldValue(t, 'lane') || 'General').toString(),
  task: (getFieldValue(t, 'Task') || `Task ${idx + 1}`).toString(),
  startDate: formatDate(getFieldValue(t, 'StartDate') || ''),
  endDate: formatDate(getFieldValue(t, 'EndDate') || ''),
  efforts: parseInt(String(getFieldValue(t, 'Effort') || '0'), 10) || 0,
  lane: normaliseLane(getFieldValue(t, 'lane') || ''),
  assignee: (getFieldValue(t, 'Assignee') || 'Unassigned').toString(),
  status: normalisePlanStatus(getFieldValue(t, 'Status') || '')
});

const filterParticipationByProject = (items: any[], projectCode: string): StaffMember[] =>
  items
    .filter((p) => projectCodesMatch(getParticipationProjectCode(p), projectCode))
    .filter((p) => getParticipationName(p))
    .map(mapParticipationToStaff);

const buildStaffByProjectCode = (participations: any[]): Map<string, StaffMember[]> => {
  const map = new Map<string, StaffMember[]>();
  participations.forEach((item) => {
    const code = getParticipationProjectCode(item);
    const name = getParticipationName(item);
    if (!code || !name) return;
    const key = code.trim().toLowerCase();
    const list = map.get(key) ?? [];
    list.push(mapParticipationToStaff(item));
    map.set(key, list);
  });
  return map;
};

const buildPlanByProjectCode = (taskItems: any[]): Map<string, PlanItem[]> => {
  const map = new Map<string, PlanItem[]>();
  taskItems.forEach((task) => {
    const code = String(getFieldValue(task, 'ProjectCode') || '').trim();
    if (!code) return;
    const key = code.toLowerCase();
    const list = map.get(key) ?? [];
    list.push(mapTaskToPlanItem(task, list.length));
    map.set(key, list);
  });
  map.forEach((items, key) => {
    map.set(key, [...items].sort((a, b) => a.slNo - b.slNo));
  });
  return map;
};

const fetchStaffForProject = async (projectCode: string): Promise<StaffMember[]> => {
  const code = (projectCode || '').trim();
  if (!code) return [];

  const sp = getSP();
  const escaped = escapeODataString(code);

  
  try {
    const byField3 = await sp.web.lists
      .getByTitle('ProjectParticipation')
      .items
      .select(...PARTICIPATION_SELECT)
      .filter(`field_3 eq '${escaped}'`)
      .top(500)();
    if (byField3.length > 0) {
      return filterParticipationByProject(byField3, code);
    }
  } catch (err) {
    console.warn('[fetchStaffForProject] field_3 OData filter failed:', err);
  }


  try {
    const allItems = await sp.web.lists
      .getByTitle('ProjectParticipation')
      .items
      .select(...PARTICIPATION_SELECT)
      .top(5000)();

    const matched = filterParticipationByProject(allItems, code);
    if (matched.length === 0 && allItems.length > 0) {
      console.warn(
        '[fetchStaffForProject] No match for project code:',
        code,
        'Available codes:',
        allItems.map((p) => getParticipationProjectCode(p)).filter(Boolean)
      );
    }
    return matched;
  } catch (error) {
    console.error('[fetchStaffForProject] Failed to load ProjectParticipation:', error);
    return [];
  }
};

const fetchPlanItemsForProject = async (projectCode: string): Promise<PlanItem[]> => {
  if (!projectCode) return [];
  const sp = getSP();
  const taskFields = ['Task', 'Sno', 'ProjectCode', 'Assignee', 'lane', 'Effort', 'Status', 'StartDate', 'EndDate'];
  const escaped = escapeODataString(projectCode);

  let items: any[] = [];
  try {
    items = await sp.web.lists
      .getByTitle('TaskDetails')
      .items
      .select(...taskFields)
      .filter(`ProjectCode eq '${escaped}'`)
      .top(5000)();
  } catch {
    items = await sp.web.lists
      .getByTitle('TaskDetails')
      .items
      .select(...taskFields)
      .top(5000)();
    items = items.filter((t) => projectCodesMatch(String(getFieldValue(t, 'ProjectCode') || ''), projectCode));
  }

  return items
    .map((t, idx) => mapTaskToPlanItem(t, idx))
    .sort((a, b) => a.slNo - b.slNo);
};

const parseBudget = (val: unknown): number => {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (typeof val === 'string') {
    const cleaned = val.replace(/[^0-9.-]+/g, '');
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  }
  return 0;
};

const normalisePriority = (raw: string): ApprovalProject['priority'] => {
  const p = (raw || 'Medium').toLowerCase();
  if (p === 'low') return 'Low';
  if (p === 'high') return 'High';
  if (p === 'critical') return 'Critical';
  return 'Medium';
};


const ClipboardCheckIcon: React.FC = () => (
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2" stroke="white" strokeWidth="1.5" strokeLinecap="round" />
    <rect x="9" y="3" width="6" height="4" rx="1" stroke="white" strokeWidth="1.5" />
    <path d="M9 12l2 2 4-4" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const CheckSvgIcon: React.FC = () => (
  <svg width="40" height="40" viewBox="0 0 40 40" fill="none" aria-hidden="true">
    <path d="M10 20l8 8 14-14" stroke="#fff" strokeWidth="3.2" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const SectionDocIcon: React.FC = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M8 4h8l4 4v12a2 2 0 01-2 2H8a2 2 0 01-2-2V6a2 2 0 012-2z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
    <path d="M16 4v4h4M10 13h6M10 17h4" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const SectionChartIcon: React.FC = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M4 19V5M4 19h16" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    <path d="M8 15V11M12 15V8M16 15v-5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const SectionTeamIcon: React.FC = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <circle cx="9" cy="8" r="3" stroke="currentColor" strokeWidth="1.8" />
    <path d="M3.5 18c0-2.6 2.4-4.5 5.5-4.5s5.5 1.9 5.5 4.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    <circle cx="17" cy="9" r="2.5" stroke="currentColor" strokeWidth="1.8" />
    <path d="M14 17.5c.7-1.6 2.3-2.7 4.5-2.7 1.3 0 2.3.4 3 .9" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const PlanTasksIcon: React.FC = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <rect x="5" y="4" width="14" height="16" rx="2" stroke="currentColor" strokeWidth="1.8" />
    <path d="M9 9h6M9 13h6M9 17h4" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const PlanEffortIcon: React.FC = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <circle cx="12" cy="12" r="8" stroke="currentColor" strokeWidth="1.8" />
    <path d="M12 8v4l3 2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const PlanFeatureIcon: React.FC = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M4 7h16M4 12h10M4 17h14" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const STEPS = ['PROJECT REVIEW', 'PLAN REVIEW'] as const;
type ReviewStep = 1 | 2 | 3; // 3 = success
const APPROVAL_SUCCESS_REDIRECT_MS = 2500;


interface ReviewDialogProps {
  project: ApprovalProject;
  onClose: () => void;
  onApprove: () => Promise<void>;
  onReject: () => void;
  onApprovalSuccess?: () => void;
}

const ReviewDialog: React.FC<ReviewDialogProps> = ({
  project,
  onClose,
  onApprove,
  onReject,
  onApprovalSuccess,
}) => {
  const css = styles as typeof styles & Record<string, string>;
  const [step, setStep] = React.useState<ReviewStep>(1);
  const [staff, setStaff] = React.useState<StaffMember[]>(project.staff);
  const [planItems, setPlanItems] = React.useState<PlanItem[]>(project.planItems);
  const [approving, setApproving] = React.useState(false);
  const [approveError, setApproveError] = React.useState<string | null>(null);

  React.useEffect(() => {
    setStep(1);
    setStaff(project.staff);
    setPlanItems(project.planItems);

    if (project.staff.length > 0 || project.planItems.length > 0) {
      return undefined;
    }

    let cancelled = false;
    void Promise.all([
      fetchStaffForProject(project.id),
      fetchPlanItemsForProject(project.id),
    ])
      .then(([staffData, planData]) => {
        if (!cancelled) {
          setStaff(staffData);
          setPlanItems(planData);
        }
      })
      .catch((err) => {
        console.error('[ReviewDialog] Failed to load project details:', err);
      });

    return () => { cancelled = true; };
  }, [project]);

  const completeApprovalSuccess = React.useCallback((): void => {
    onClose();
    onApprovalSuccess?.();
  }, [onClose, onApprovalSuccess]);

  React.useEffect(() => {
    if (step !== 3) return undefined;
    const timer = window.setTimeout(() => {
      completeApprovalSuccess();
    }, APPROVAL_SUCCESS_REDIRECT_MS);
    return () => { window.clearTimeout(timer); };
  }, [step, completeApprovalSuccess]);

  const initials = (name: string): string =>
    name.split(' ').map(p => p[0]).join('').slice(0, 2).toUpperCase();

  const getLaneCls = (lane: PlanItem['lane']): string => {
    if (lane === 'testing') return styles.testing;
    if (lane === 'demo')    return styles.demo;
    return styles.dev;
  };

  const getStatusCls = (status: PlanItem['status']): string => {
    if (status === 'In Progress') return styles.inProgress;
    if (status === 'Done')        return styles.done;
    return styles.notStarted;
  };

  const totalEfforts = planItems.reduce((s, i) => s + i.efforts, 0);
  const uniqueFeatures = planItems.reduce<string[]>((acc, i) => {
    if (acc.indexOf(i.feature) === -1) acc.push(i.feature);
    return acc;
  }, []).length;

  const handleExportPlan = (): void => {
    const headers = [
      'Sl.No',
      'Feature',
      'Task',
      'Type',
      'Start Date',
      'End Date',
      'Effort (hrs)',
      'Assignee',
      'Status',
    ];
    const rows = planItems.map((item, idx) => [
      item.slNo || idx + 1,
      item.feature,
      item.task,
      item.lane,
      item.startDate,
      item.endDate,
      item.efforts,
      item.assignee,
      item.status,
    ]);
    const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'Project Plan');
    const safeName = (project.name || project.id || 'Project Plan').replace(/[\\/:*?"<>|]+/g, '-').trim();
    XLSX.writeFile(workbook, `${safeName} - Project Plan.xlsx`);
  };

  const priorityClass = (): string => {
    const p = project.priority.toLowerCase();
    if (p === 'high')     return styles.high;
    if (p === 'medium')   return styles.medium;
    if (p === 'low')      return styles.low;
    return styles.critical;
  };

  const handleApprove = async (): Promise<void> => {
    setApproving(true);
    setApproveError(null);
    try {
      await onApprove();
      setStep(3);
    } catch (error) {
      console.error('[ReviewDialog] Failed to approve project:', error);
      setApproveError('Could not update isManagerApproved. Please try again.');
    } finally {
      setApproving(false);
    }
  };

  return (
    <div className={styles.overlay} onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className={styles.dialogShell}>
        <div className={styles.dialogGlow} aria-hidden="true" />
        <div className={styles.dialog} role="dialog" aria-modal="true" aria-label="Project approval review">

        {/* Header */}
        <div className={styles.dialogHeader}>
          <div className={styles.dialogHeaderLeft}>
            <div className={styles.dialogLogo}>
              <ClipboardCheckIcon />
            </div>
            <div>
              <h2 className={styles.dialogTitle}>Approval Review Wizard</h2>
              <p className={styles.dialogSubtitle}>
                {project.id} · {project.name}
              </p>
            </div>
          </div>
          <button type="button" className={styles.dialogClose} aria-label="Close" onClick={onClose}>×</button>
        </div>

        {/* Stepper */}
        {step < 3 && (
          <div className={styles.stepper}>
            {STEPS.map((label, idx) => {
              const stepNum = idx + 1;
              const isActive = stepNum === step;
              const isDone   = stepNum < step;
              const isLast   = idx === STEPS.length - 1;
              const connectorClass = isDone
                ? styles.connectorDone
                : (isActive && !isLast ? styles.connectorActive : '');
              return (
                <div key={label} className={styles.stepItem}>
                  {!isLast && (
                    <div className={`${styles.stepConnector} ${connectorClass}`} />
                  )}
                  <div className={`${styles.stepCircle} ${isActive ? styles.active : ''} ${isDone ? styles.done : ''}`}>
                    {isDone ? '✓' : stepNum}
                  </div>
                  <span className={`${styles.stepLabel} ${isActive ? styles.active : ''} ${isDone ? styles.done : ''}`}>
                    {label}
                  </span>
                </div>
              );
            })}
          </div>
        )}

        {/* Body */}
        <div className={styles.dialogBody}>

         
          {step === 1 && (
            <>
              <div className={styles.reviewGrid}>
            
                <div className={styles.reviewSection}>
                  <h3 className={styles.reviewSectionTitle}>
                    <span className={styles.reviewSectionIcon}><SectionDocIcon /></span>
                    Project Details
                  </h3>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Project Name</span>
                    <span className={styles.reviewFieldValue}>{project.name}</span>
                  </div>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Client</span>
                    <span className={styles.reviewFieldValue}>{project.client}</span>
                  </div>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Budget</span>
                    <span className={styles.reviewFieldValue}>{project.budget}</span>
                  </div>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Timeline</span>
                    <span className={styles.reviewFieldValue}>{project.startDate} → {project.endDate}</span>
                  </div>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Priority</span>
                    <span className={`${styles.priorityTag} ${priorityClass()}`}>{project.priority}</span>
                  </div>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Required Skills</span>
                    <div className={styles.skillsWrap}>
                      {project.skills.map(s => (
                        <span key={s} className={styles.skillTag}>{s}</span>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Brief */}
                <div className={styles.reviewSection}>
                  <h3 className={styles.reviewSectionTitle}>
                    <span className={styles.reviewSectionIcon}><SectionChartIcon /></span>
                    Project Summary
                  </h3>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Scope</span>
                    <span className={styles.reviewFieldValue}>{project.subtitle}</span>
                  </div>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Project ID</span>
                    <span className={styles.reviewFieldValue}>{project.id}</span>
                  </div>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Team Size</span>
                    <span className={styles.reviewFieldValue}>
                      {staff.length} member{staff.length === 1 ? '' : 's'}
                    </span>
                  </div>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Plan Tasks</span>
                    <span className={styles.reviewFieldValue}>
                      {planItems.length} tasks across {uniqueFeatures} feature{uniqueFeatures === 1 ? '' : 's'}
                    </span>
                  </div>
                  <div className={styles.reviewField}>
                    <span className={styles.reviewFieldLabel}>Total Effort</span>
                    <span className={styles.reviewFieldValue}>{totalEfforts} hours</span>
                  </div>
                </div>
              </div>

              <div className={styles.staffSection}>
                <h3 className={styles.reviewSectionTitle}>
                  <span className={styles.reviewSectionIcon}><SectionTeamIcon /></span>
                 
                  Recommended Resource Allocation

                </h3>
                {staff.length === 0 ? (
                  <div className={css.staffEmpty}>
                    No staff found in <strong>ProjectParticipation</strong> for project code <code>{project.id}</code>.
                  </div>
                ) : (
                  <div className={styles.staffGrid}>
                    {staff.map((member, idx) => (
                      <div key={`${member.name}-${idx}`} className={styles.staffCard}>
                        <div className={styles.staffCardTop}>
                          <span className={styles.roleBadge}>{member.role}</span>
                          <span className={styles.matchScore}>{member.match}% match</span>
                        </div>
                        <div className={styles.staffProfile}>
                          <div className={styles.staffAvatar}>{member.initials}</div>
                          <div>
                            <p className={styles.staffName}>{member.name}</p>
                            <p className={styles.staffDept}>{member.department || 'Project Team'}</p>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </>
          )}


          {step === 2 && (
            <>
        
              <div className={styles.planSummaryRow}>
                <div className={styles.planKpi}>
                  <span className={styles.planKpiIcon}><PlanTasksIcon /></span>
                  <div>
                    <span className={styles.planKpiLabel}>Total Tasks</span>
                    <span className={styles.planKpiValue}>{planItems.length}</span>
                  </div>
                </div>
                <div className={styles.planKpi}>
                  <span className={`${styles.planKpiIcon} ${styles.planKpiIconGreen}`}><PlanEffortIcon /></span>
                  <div>
                    <span className={styles.planKpiLabel}>Total Effort</span>
                    <span className={styles.planKpiValue}>{totalEfforts}h</span>
                  </div>
                </div>
                <div className={styles.planKpi}>
                  <span className={`${styles.planKpiIcon} ${styles.planKpiIconBlue}`}><PlanFeatureIcon /></span>
                  <div>
                    <span className={styles.planKpiLabel}>Features</span>
                    <span className={styles.planKpiValue}>{uniqueFeatures}</span>
                  </div>
                </div>
                <div className={styles.planKpi}>
                  <span className={`${styles.planKpiIcon} ${styles.planKpiIconPurple}`}><SectionTeamIcon /></span>
                  <div>
                    <span className={styles.planKpiLabel}>Team Size</span>
                    <span className={styles.planKpiValue}>{staff.length}</span>
                  </div>
                </div>
              </div>

              <div className={css.planReviewToolbar}>
                <span className={css.planReviewToolbarLabel}>Project Plan</span>
                <button
                  type="button"
                  className={css.planExportBtn}
                  onClick={handleExportPlan}
                  disabled={planItems.length === 0}
                >
                  Export Excel
                </button>
              </div>

          
              <div className={styles.planReviewWrap}>
                <table className={styles.planTable}>
                  <thead>
                    <tr>
                      <th>#</th>
                      <th>Feature</th>
                      <th>Task</th>
                      <th>Type</th>
                      <th>Start Date</th>
                      <th>End Date</th>
                      <th className={styles.centered}>Efforts (h)</th>
                      <th>Assignee</th>
                      <th className={styles.centered}>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {planItems.map((item, idx) => (
                      <tr key={idx}>
                        <td className={styles.planIndexCell}>{idx + 1}</td>
                        <td className={styles.planFeatureCell}>{item.feature}</td>
                        <td className={styles.planTaskCell}>{item.task}</td>
                        <td>
                          <span className={`${styles.planLaneBadge} ${getLaneCls(item.lane)}`}>
                            {item.lane}
                          </span>
                        </td>
                        <td className={styles.planDateCell}>{item.startDate}</td>
                        <td className={styles.planDateCell}>{item.endDate}</td>
                        <td className={styles.planEffortsCell}>{item.efforts}</td>
                        <td className={styles.planAssigneeCell}>
                          <div className={styles.planAvatarWrap}>
                            <div className={styles.planAvatar}>{initials(item.assignee)}</div>
                            {item.assignee}
                          </div>
                        </td>
                        <td>
                          <span className={`${styles.planStatusBadge} ${getStatusCls(item.status)}`}>
                            {item.status}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}

          {step === 3 && (
            <div className={styles.successPanel}>
              <div className={styles.successHero}>
                <div className={styles.successHalo} />
                <div className={styles.successIcon}>
                  <CheckSvgIcon />
                </div>
              </div>
              <h2 className={styles.successTitle}>Project Approved Successfully!</h2>
              <p className={styles.successDesc}>
                <strong>{project.name}</strong> has been approved and is now authorized.
                Returning you to the dashboard to view the updated project list.
              </p>
              <div className={styles.successMeta}>
                <span className={styles.successMetaTag}>✓ {project.id}</span>
                <span className={styles.successMetaTag}>✓ {staff.length} staff allocated</span>
                <span className={styles.successMetaTag}>✓ Plan authorized</span>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        {step < 3 ? (
          <div className={styles.dialogFooter}>
            <button
              type="button"
              className={styles.btnBack}
              onClick={() => step > 1 ? setStep((step - 1) as ReviewStep) : onClose()}
            >
              {step === 1 ? 'Cancel' : '← Back'}
            </button>
            <div className={styles.footerRight}>
              {step === 2 && (
                <button type="button" className={styles.btnReject} onClick={onReject}>
                  ✕ Reject Project
                </button>
              )}
              {step === 1 && (
                <button type="button" className={styles.btnNext} onClick={() => setStep(2)}>
                  Review Plan
                </button>
              )}
              {step === 2 && (
                <button
                  type="button"
                  className={styles.btnApprove}
                  disabled={approving}
                  onClick={() => { handleApprove().catch(console.error); }}
                >
                  {approving ? 'Approving…' : '✓ Approve Project'}
                </button>
              )}
              {approveError && (
                <p className={css.approveError} role="alert">{approveError}</p>
              )}
            </div>
          </div>
        ) : (
          <div className={`${styles.dialogFooter} ${css.dialogFooterSuccess}`}>
            <button
              type="button"
              className={styles.btnApprove}
              onClick={completeApprovalSuccess}
            >
              Go to Dashboard
            </button>
          </div>
        )}
        </div>
      </div>
    </div>
  );
};



export interface IApprovalsProps {
  onNavigate?: (view: string) => void;
  onApprovalSuccess?: () => void;
}

const PendingIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.8" />
    <path d="M12 7v5l3 2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const ApprovedIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.8" />
    <path d="M8 12l2.5 2.5L16 9" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const QueueIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <rect x="4" y="5" width="16" height="14" rx="2" stroke="currentColor" strokeWidth="1.8" />
    <path d="M8 9h8M8 13h5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const getPriorityClass = (priority: string, css: Record<string, string>): string => {
  const p = (priority || '').toLowerCase();
  if (p === 'critical') return css.priorityCritical;
  if (p === 'high') return css.priorityHigh;
  if (p === 'medium') return css.priorityMedium;
  return css.priorityLow;
};

const Approvals: React.FC<IApprovalsProps> = ({ onNavigate, onApprovalSuccess }) => {
  const css = styles as typeof styles & Record<string, string>;
  const { displayName, initials } = useUser();

  const [projects, setProjects] = React.useState<ApprovalProject[]>([]);
  const [approvedCount, setApprovedCount] = React.useState(0);
  const [loading, setLoading] = React.useState(true);
  const [fetchError, setFetchError] = React.useState<string | null>(null);
  const [reviewProject, setReviewProject] = React.useState<ApprovalProject | null>(null);
  const [searchTerm, setSearchTerm] = React.useState('');

  const pending  = projects.filter(p => p.status === 'Pending').length;
  const approved = approvedCount;
  const total    = pending + approved;

  const filteredProjects = React.useMemo(() => {
    let list = projects;
    if (searchTerm.trim()) {
      const q = searchTerm.toLowerCase();
      list = projects.filter(
        (p) =>
          p.id.toLowerCase().includes(q) ||
          p.name.toLowerCase().includes(q) ||
          p.client.toLowerCase().includes(q) ||
          p.priority.toLowerCase().includes(q)
      );
    }
    return sortProjectsByModifiedDesc(list);
  }, [projects, searchTerm]);

  const fetchApprovals = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      setFetchError(null);
      const sp = getSP();

      const projectFields = [
        'Id', 'ProjectCode', 'ProjectName', 'Client', 'StartDate', 'EndDate',
        'Budget', 'RequiredSkills', 'Description', 'Priority', 'isManagerApproved', 'Modified'
      ];

      const loadProjects = async (): Promise<{ spItems: any[]; approvedTotal: number }> => {
        const fetchPending = (withOrder: boolean): Promise<any[]> => {
          let query = sp.web.lists
            .getByTitle('Project')
            .items
            .select(...projectFields)
            .filter('isManagerApproved eq false');
          if (withOrder) {
            query = query.orderBy('Modified', false);
          }
          return query.top(5000)();
        };

        try {
          let pendingItems: any[] = [];
          try {
            pendingItems = await fetchPending(true);
          } catch {
            pendingItems = await fetchPending(false);
          }

          const approvedItems = await sp.web.lists
            .getByTitle('Project')
            .items
            .select('Id')
            .filter('isManagerApproved eq true')
            .top(5000)();

          return {
            spItems: pendingItems.filter(isManagerApprovedFalse),
            approvedTotal: approvedItems.length,
          };
        } catch {
          let allProjects: any[] = [];
          try {
            allProjects = await sp.web.lists
              .getByTitle('Project')
              .items
              .select(...projectFields)
              .orderBy('Modified', false)
              .top(5000)();
          } catch {
            allProjects = await sp.web.lists
              .getByTitle('Project')
              .items
              .select(...projectFields)
              .top(5000)();
          }
          return {
            spItems: allProjects.filter(isManagerApprovedFalse),
            approvedTotal: allProjects.filter((item) => getManagerApprovedValue(item) === true).length,
          };
        }
      };

      const [projectResult, participations, taskItems] = await Promise.all([
        loadProjects(),
        sp.web.lists
          .getByTitle('ProjectParticipation')
          .items
          .select(...PARTICIPATION_SELECT)
          .top(5000)()
          .catch(() => [] as any[]),
        sp.web.lists
          .getByTitle('TaskDetails')
          .items
          .select('Task', 'Sno', 'ProjectCode', 'Assignee', 'lane', 'Effort', 'Status', 'StartDate', 'EndDate')
          .top(5000)()
          .catch(() => [] as any[]),
      ]);

      const spItems = projectResult.spItems;
      setApprovedCount(projectResult.approvedTotal);

      const staffByProject = buildStaffByProjectCode(participations);
      const planByProject = buildPlanByProjectCode(taskItems);

      const mapped: ApprovalProject[] = spItems.map((item) => {
        const projectCode = (getFieldValue(item, 'ProjectCode') || '').toString();
        const projectKey = projectCode.trim().toLowerCase();
        const spItemId    = Number(item.Id ?? getFieldValue(item, 'Id')) || 0;
        const budgetNum   = parseBudget(getFieldValue(item, 'Budget'));
        const skillsRaw   = getFieldValue(item, 'RequiredSkills') || '';
        const skills      = skillsRaw
          ? skillsRaw.split(',').map((s: string) => s.trim()).filter(Boolean)
          : [];

        const staff = staffByProject.get(projectKey) ?? [];
        const planItems = planByProject.get(projectKey) ?? [];
        const members = staff.map((member) => member.name);
        const visibleTeam = members.slice(0, 3).map(getInitials);

        const modifiedVal = getFieldValue(item, 'Modified') || item.Modified || '';

        return {
          spItemId,
          id: projectCode || `PRJ-${spItemId}`,
          name: (getFieldValue(item, 'ProjectName') || 'Untitled Project').toString(),
          subtitle: (getFieldValue(item, 'Description') || '').toString(),
          client: (getFieldValue(item, 'Client') || '—').toString(),
          budget: `$${budgetNum.toLocaleString()}`,
          startDate: formatDate(getFieldValue(item, 'StartDate') || ''),
          endDate: formatDate(getFieldValue(item, 'EndDate') || ''),
          priority: normalisePriority(getFieldValue(item, 'Priority') || ''),
          skills,
          team: visibleTeam,
          teamExtra: Math.max(0, members.length - 3),
          status: 'Pending',
          staff,
          planItems,
          modified: String(modifiedVal)
        };
      });

      setProjects(sortProjectsByModifiedDesc(mapped));
    } catch (error) {
      console.error('Approvals: Error fetching projects from SharePoint:', error);
      setProjects([]);
      setApprovedCount(0);
      setFetchError('Could not load projects from SharePoint. Please refresh and try again.');
    } finally {
      setLoading(false);
    }
  }, []);

  React.useEffect(() => {
    fetchApprovals().catch(console.error);
  }, [fetchApprovals]);

  const handleApprove = async (): Promise<void> => {
    if (!reviewProject) {
      throw new Error('No project selected for approval');
    }
    if (!reviewProject.spItemId || reviewProject.spItemId <= 0) {
      throw new Error(`Missing SharePoint item ID for project ${reviewProject.id}`);
    }

    const sp = getSP();
    await sp.web.lists
      .getByTitle('Project')
      .items
      .getById(reviewProject.spItemId)
      .update({ isManagerApproved: true });

    setProjects((prev) => prev.filter((p) => p.id !== reviewProject.id));
    setApprovedCount((prev) => prev + 1);
  };

  const handleReject = (): void => {
    if (!reviewProject) return;
    setProjects(prev => prev.filter(p => p.id !== reviewProject.id));
    setReviewProject(null);
  };

  return (
    <div className={css.page}>
      <div className={css.pageMesh} aria-hidden="true" />
      <div className={css.pageOrb} aria-hidden="true" />
      <div className={css.pageOrb2} aria-hidden="true" />

      <NavBar
        activeItem="approvals"
        onNavigate={onNavigate}
        pendingApprovalsCount={projects.length}
      />

      <main className={css.main}>
        <header className={css.topBar}>
          <div className={css.topBarLeft}>
            <h1 className={css.portalTitle}>Manager Approval Portal</h1>
            <p className={css.portalSubtitle}>
              Review submitted project intake requests and authorize provisioning.
            </p>
          </div>
          <div className={css.topBarRight}>
            <div className={css.profileCard}>
              <span className={css.profileAvatar}>{initials}</span>
              <div className={css.profileText}>
                <strong>{displayName}</strong>
                <span>Project Manager</span>
              </div>
              <svg className={css.profileChevron} viewBox="0 0 24 24" aria-hidden="true" fill="none">
                <path d="M6 9l6 6 6-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </div>
          </div>
        </header>

        <section className={css.statsRow} aria-label="Approval summary">
          <article className={`${css.statCard} ${css.statPending}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><PendingIcon /></span>
            </div>
            <p className={css.statLabel}>Pending Approval</p>
            <p className={css.statValue}>{pending}</p>
            <span className={`${css.statTrend} ${css.statTrendPurple}`}>Requires Action</span>
            <div className={`${css.statWatermark} ${css.statWatermarkPurple}`} aria-hidden="true">
              <PendingIcon />
            </div>
          </article>
          <article className={`${css.statCard} ${css.statApproved}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><ApprovedIcon /></span>
            </div>
            <p className={css.statLabel}>Approved</p>
            <p className={css.statValue}>{approved}</p>
            <span className={`${css.statTrend} ${css.statTrendGreen}`}>Authorized</span>
            <div className={`${css.statWatermark} ${css.statWatermarkGreen}`} aria-hidden="true">
              <ApprovedIcon />
            </div>
          </article>
          <article className={`${css.statCard} ${css.statTotal}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><QueueIcon /></span>
            </div>
            <p className={css.statLabel}>Total Submitted</p>
            <p className={css.statValue}>{total}</p>
            <span className={`${css.statTrend} ${css.statTrendBlue}`}>This Cycle</span>
            <div className={`${css.statWatermark} ${css.statWatermarkBlue}`} aria-hidden="true">
              <QueueIcon />
            </div>
          </article>
        </section>

        <section className={css.card}>
          <div className={css.cardGlow} aria-hidden="true" />
          <div className={css.cardInner}>
            <div className={css.cardHead}>
              <div className={css.cardHeadText}>
                <span className={css.cardKicker}>Approval Queue</span>
                <h3>Pending Project Submissions</h3>
                <p>Review intake details, staffing recommendations, and project plans before authorization.</p>
              </div>
              <div className={css.searchBox}>
                <svg className={css.searchIcon} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                  <circle cx="11" cy="11" r="8" />
                  <path d="M21 21l-4.35-4.35" />
                </svg>
                <input
                  type="text"
                  className={css.searchInput}
                  placeholder="Search submissions..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  aria-label="Search approval submissions"
                />
              </div>
            </div>

            {loading ? (
              <div className={css.loadingWrap} aria-live="polite">
                <div className={css.spinner} />
                <span>Loading pending approvals…</span>
              </div>
            ) : fetchError ? (
              <div className={css.emptyWrap} aria-live="polite">
                <p className={css.emptyTitle}>Unable to load approvals</p>
                <p className={css.emptyDesc}>{fetchError}</p>
              </div>
            ) : projects.length === 0 ? (
              <div className={css.emptyWrap} aria-live="polite">
                <p className={css.emptyTitle}>No pending approvals</p>
              </div>
            ) : filteredProjects.length === 0 ? (
              <div className={css.emptyWrap} aria-live="polite">
                <p className={css.emptyTitle}>No matching submissions</p>
                <p className={css.emptyDesc}>Try a different search term to find pending approvals.</p>
              </div>
            ) : (
            <div className={css.tableWrap}>
              <table className={css.table}>
                <thead>
                  <tr>
                    <th>Project ID</th>
                    <th>Project Name</th>
                    <th>Client</th>
                    <th>Priority</th>
                    <th>Start Date</th>
                    <th>End Date</th>
                    <th className={css.budgetCol}>Budget</th>
                    <th>Status</th>
                    <th className={css.colAction}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredProjects.map((project) => (
                    <tr key={project.id} className={css.projectRow}>
                      <td className={css.idCell}>
                        <span className={css.idPill}>{project.id}</span>
                      </td>
                      <td className={css.nameCell}>
                        <strong>{project.name}</strong>
                      </td>
                      <td className={css.dateCell}>{project.client}</td>
                      <td>
                        <span className={`${css.priorityPill} ${getPriorityClass(project.priority, css)}`}>
                          {project.priority}
                        </span>
                      </td>
                      <td className={css.dateCell}>{project.startDate || '—'}</td>
                      <td className={css.dateCell}>{project.endDate || '—'}</td>
                      <td className={`${css.amountCell} ${css.budgetCol}`}>
                        <span className={css.amountVal}>{project.budget}</span>
                        <small>USD</small>
                      </td>
                      <td>
                        <span className={`${css.statusPill} ${
                          project.status === 'Approved' ? css.statusApproved
                            : project.status === 'Rejected' ? css.statusRejected
                            : css.statusPending
                        }`}>
                          <span className={css.statusDot} />
                          {project.status}
                        </span>
                      </td>
                      <td className={css.colAction}>
                        <button
                          type="button"
                          className={css.actionBtn}
                          disabled={project.status !== 'Pending'}
                          onClick={() => setReviewProject(project)}
                        >
                          Review &amp; Act
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            )}

            {!loading && !fetchError && projects.length > 0 && (
              <div className={css.cardFoot}>
                <span>
                  Showing <strong>{filteredProjects.length}</strong> of <strong>{projects.length}</strong> submissions
                </span>
                <span className={css.pendingHint}>
                  {pending} pending review
                </span>
              </div>
            )}
          </div>
        </section>
      </main>

     
      {reviewProject && (
        <ReviewDialog
          project={reviewProject}
          onClose={() => setReviewProject(null)}
          onApprove={() => handleApprove()}
          onReject={handleReject}
          onApprovalSuccess={onApprovalSuccess}
        />
      )}
    </div>
  );
};

export default Approvals;
