import * as React from 'react';
import styles from './dashboard.module.scss';
import NavBar from '../NavBar/navbar';
import { useUser } from '../UserContext/UserContext';
import { getSP } from '../../../../service/initservice';


export interface ProjectRow {
  id: string;
  name: string;
  subtitle: string;
  startDate: string;
  endDate: string;
  budget: string;
  team: string[];
  teamExtra: number;

  modified?: string;
}

export interface IDashboardProps {
  onCreateProject: () => void;
  onOpenProject: (project: ProjectRow) => void;
  onNavigate?: (view: string) => void;
  refreshKey?: number;
  approvalsRefreshKey?: number;
}

const AVATAR_COLORS = ['#dbeafe', '#e0e7ff', '#fce7f3', '#d1fae5', '#fef3c7'];
const AVATAR_TEXT = ['#6d28d9', '#7c3aed', '#be185d', '#047857', '#b45309'];
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

const normalizeApprovedFlag = (raw: unknown): boolean | null => {
  if (raw === undefined || raw === null || raw === '') return null;
  if (typeof raw === 'boolean') return raw;
  if (typeof raw === 'number') return raw === 1 ? true : raw === 0 ? false : null;
  const text = String(raw).trim().toLowerCase();
  if (text === 'true' || text === '1' || text === 'yes') return true;
  if (text === 'false' || text === '0' || text === 'no') return false;
  return null;
};

const getManagerApprovedValue = (item: any): boolean | null => {
  if (!item) return null;
  const candidates: unknown[] = [
    item.isManagerApproved,
    item.IsManagerApproved,
    getFieldValue(item, 'isManagerApproved'),
    getFieldValue(item, 'IsManagerApproved'),
  ];

  for (const key of Object.keys(item)) {
    if (key.toLowerCase().includes('managerapproved')) {
      candidates.push(item[key]);
    }
  }

  for (const raw of candidates) {
    const parsed = normalizeApprovedFlag(raw);
    if (parsed !== null) return parsed;
  }

  return null;
};

const isManagerApprovedTrue = (item: any): boolean =>
  getManagerApprovedValue(item) === true;
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

const parseBudget = (val: unknown): number => {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (typeof val === 'string') {
    const cleaned = val.replace(/[^0-9.-]+/g, '');
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  }
  return 0;
};

const parseModifiedTime = (value?: string): number => {
  if (!value) return 0;
  const time = new Date(value).getTime();
  return isNaN(time) ? 0 : time;
};


const sortProjectsByModifiedDesc = (items: ProjectRow[]): ProjectRow[] =>
  [...items].sort((a, b) => {
    const diff = parseModifiedTime(b.modified) - parseModifiedTime(a.modified);
    if (diff !== 0) return diff;
    return b.id.localeCompare(a.id, undefined, { numeric: true, sensitivity: 'base' });
  });


const FolderIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <path d="M3 7.5A2.5 2.5 0 015.5 5H9l2 2h7.5A2.5 2.5 0 0119 9.5v9A2.5 2.5 0 0116.5 21h-11A2.5 2.5 0 013 18.5v-11z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
  </svg>
);

const DollarIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.8" />
    <path d="M12 7v10M9.5 9.5c0-1.2 1.1-2 2.5-2s2.5.8 2.5 2-1.1 2-2.5 2.2-2.5 2.8-2.5 2.8-1.4 0-2.5.8-2.5 2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const WalletIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <path d="M4 7.5A2.5 2.5 0 016.5 5H17a2 2 0 012 2v2H6.5A2.5 2.5 0 004 11.5v9A2.5 2.5 0 006.5 23H19a2 2 0 002-2v-9a2 2 0 00-2-2H8" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
    <circle cx="16.5" cy="14.5" r="1.5" fill="currentColor" />
  </svg>
);

const ClockIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.8" />
    <path d="M12 8v4l3 2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const formatBudgetCompact = (amount: number): string => {
  if (amount >= 1_000_000) return `$${(amount / 1_000_000).toFixed(1)}M`;
  if (amount >= 1_000) return `$${(amount / 1_000).toFixed(1)}K`;
  return `$${amount.toLocaleString()}`;
};

const Dashboard: React.FC<IDashboardProps> = ({
  onCreateProject,
  onOpenProject,
  onNavigate,
  refreshKey = 0,
  approvalsRefreshKey = 0
}) => {
  const css = styles as typeof styles & Record<string, string>;
  const { displayName, initials } = useUser();

  const [projects, setProjects] = React.useState<ProjectRow[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);


  const projectsToRender = React.useMemo(
    () => sortProjectsByModifiedDesc(projects),
    [projects]
  );


  const totalBudget = React.useMemo(
    () => projectsToRender.reduce((sum, p) => sum + parseBudget(p.budget), 0),
    [projectsToRender]
  );

  const avgDurationMonths = React.useMemo(() => {
    let totalMonths = 0;
    let valid = 0;
    projectsToRender.forEach((p) => {
      const start = new Date(p.startDate);
      const end = new Date(p.endDate);
      if (!isNaN(start.getTime()) && !isNaN(end.getTime())) {
        totalMonths += Math.abs(end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24 * 30.4);
        valid++;
      }
    });
    return valid > 0 ? (totalMonths / valid).toFixed(1) : '0';
  }, [projectsToRender]);


  const fetchProjects = React.useCallback(async (): Promise<void> => {
    try {
      setLoading(true);
      const sp = getSP();

      const projectFields = [
        'Id',
        'ProjectCode',
        'ProjectName',
        'Client',
        'StartDate',
        'EndDate',
        'Budget',
        'RequiredSkills',
        'Description',
        'Modified',
        'isManagerApproved',
      ];

      let allItems: any[] = [];
      try {
        allItems = await sp.web.lists
          .getByTitle('Project')
          .items
          .select(...projectFields)
          .orderBy('Modified', false)
          .top(5000)();
      } catch {
        allItems = await sp.web.lists
          .getByTitle('Project')
          .items
          .select(...projectFields)
          .top(5000)();
      }

      const spItems = allItems.filter(isManagerApprovedTrue);
      let participations: any[] = [];
      try {
        participations = await sp.web.lists
          .getByTitle('ProjectParticipation')
          .items
          .select('field_2', 'field_3')
          .top(5000)();
      } catch {

      }

      const mapped: ProjectRow[] = spItems.map((item) => {
        const projectCode = getFieldValue(item, 'ProjectCode') || '';
        const budgetVal = getFieldValue(item, 'Budget');
        const budgetNum = parseBudget(budgetVal);
        const budgetStr = `$${budgetNum.toLocaleString()}`;
        const nameVal = getFieldValue(item, 'ProjectName') || 'Untitled Project';
        const descVal = getFieldValue(item, 'Description') || getFieldValue(item, 'Client') || '';
        const startVal = getFieldValue(item, 'StartDate') || '';
        const endVal = getFieldValue(item, 'EndDate') || '';
        const spId = getFieldValue(item, 'Id') || '';


        const members = participations
          .filter((p) => p.field_3 === projectCode && p.field_2)
          .map((p) => String(p.field_2));

        const visibleTeam = members.slice(0, 3).map(getInitials);
        const extraCount = Math.max(0, members.length - 3);

        const modifiedVal = getFieldValue(item, 'Modified') || item.Modified || '';

        return {
          id: projectCode || `PRJ-${spId}`,
          name: nameVal,
          subtitle: descVal,
          startDate: formatDate(startVal),
          endDate: formatDate(endVal),
          budget: budgetStr,
          team: visibleTeam,
          teamExtra: extraCount,
          modified: String(modifiedVal)
        } as ProjectRow;
      });

      setProjects(sortProjectsByModifiedDesc(mapped));
    } catch (error) {
      console.error('Dashboard: Error fetching projects from SharePoint:', error);
      setProjects([]);
    } finally {
      setLoading(false);
    }
  }, []);

  React.useEffect(() => {
    fetchProjects().catch(console.error);
  }, [fetchProjects]);

  React.useEffect(() => {
    if (refreshKey < 1) return;
    fetchProjects().catch(console.error);
  }, [refreshKey, fetchProjects]);

  const projectCount = projectsToRender.length;
  const projectsOnScheduleLabel =
    projectCount === 1 ? '1 Project On Schedule' : `${projectCount} Projects On Schedule`;


  return (
    <div className={css.page}>
      <div className={css.pageMesh} aria-hidden="true" />
      <div className={css.pageOrb} aria-hidden="true" />
      <div className={css.pageOrb2} aria-hidden="true" />

      <NavBar activeItem="dashboard" onNavigate={onNavigate} refreshKey={approvalsRefreshKey} />

      <main className={css.main}>

        <header className={css.topBar}>
          <div className={css.topBarLeft}>
            <h1 className={css.portalTitle}>Project Management Portal</h1>
            <p className={css.portalSubtitle}>
              Track project progress, timelines, budgets and resource assignments.
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

        {/* ── Stats row ── */}
        <section className={css.statsRow} aria-label="Workspace summary">
          <article className={`${css.statCard} ${css.statProjects}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><FolderIcon /></span>
              <button type="button" className={css.statMenuBtn} aria-label="Projects options">⋯</button>
            </div>
            <p className={css.statLabel}>Projects in Progress</p>
            <p className={css.statValue}>{projectCount}</p>
            <span className={`${css.statTrend} ${css.statTrendPurple}`}>{projectsOnScheduleLabel}</span>
            <div className={`${css.statWatermark} ${css.statWatermarkPurple}`} aria-hidden="true">
              <FolderIcon />
            </div>
          </article>
          <article className={`${css.statCard} ${css.statBudget}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><WalletIcon /></span>
              <button type="button" className={css.statMenuBtn} aria-label="Budget options">⋯</button>
            </div>
            <p className={css.statLabel}>Approved Budget</p>
            <p className={css.statValue}>{formatBudgetCompact(totalBudget)}</p>
            <span className={`${css.statTrend} ${css.statTrendBlue}`}>FY26 Budget Approved</span>
            <div className={`${css.statWatermark} ${css.statWatermarkBlue}`} aria-hidden="true">
              <DollarIcon />
            </div>
          </article>
          <article className={`${css.statCard} ${css.statDuration}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><ClockIcon /></span>
              <button type="button" className={css.statMenuBtn} aria-label="Timeline options">⋯</button>
            </div>
            <p className={css.statLabel}>Average Timeline</p>
            <p className={css.statValue}>{avgDurationMonths} Mo</p>
            <span className={`${css.statTrend} ${css.statTrendGreen}`}>Portfolio Average</span>
            <div className={`${css.statWatermark} ${css.statWatermarkGreen}`} aria-hidden="true">
              <ClockIcon />
            </div>
          </article>
        </section>

      
        <section className={css.card}>
          <div className={css.cardGlow} aria-hidden="true" />
          <div className={css.cardInner}>
            <div className={css.cardHead}>
              <div className={css.cardHeadText}>
                <span className={css.cardKicker}>Project Portfolio</span>
                <h3>Projects Under Execution</h3>
                <p>Monitor project schedules, budgets, resources and key delivery milestones.</p>
              </div>
              <button type="button" className={css.createBtn} onClick={onCreateProject}>
                <span className={css.createBtnIcon} aria-hidden="true">+</span>
                New Project
              </button>
            </div>

            {loading ? (
              <div className={css.loadingWrap} aria-live="polite">
                <div className={css.spinner} />
                <span>Loading projects…</span>
              </div>
            ) : (
              <div className={css.tableWrap}>
                <table className={css.table}>
                  <thead>
                    <tr>
                      <th>Project ID</th>
                      <th>Project Name</th>
                      <th>Start Date</th>
                      <th>End Date</th>
                      <th className={css.budgetCol}>Budgeted Amount</th>
                      <th className={css.teamCol}>Team Members</th>
                      <th className={css.colAction} aria-hidden="true" />
                    </tr>
                  </thead>
                  <tbody>
                    {projectsToRender.map((project) => (
                      <tr
                        key={project.id}
                        className={css.projectRow}
                        onClick={() => onOpenProject(project)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter' || e.key === ' ') {
                            e.preventDefault();
                            onOpenProject(project);
                          }
                        }}
                        tabIndex={0}
                        role="button"
                        aria-label={`Open ${project.name}`}
                      >
                        <td className={css.idCell}>
                          <span className={css.idPill}>{project.id}</span>
                        </td>
                        <td className={css.nameCell}>
                          <strong>{project.name}</strong>
                        </td>
                        <td className={css.dateCell}>{project.startDate}</td>
                        <td className={css.dateCell}>{project.endDate}</td>
                        <td className={`${css.amountCell} ${css.budgetCol}`}>
                          <span className={css.amountVal}>{project.budget}</span>
                          <small>USD</small>
                        </td>
                        <td className={css.teamCol}>
                          <div className={css.teamRow}>
                            {(project.team || []).length === 0 ? (
                              <span className={css.teamEmpty}>No team</span>
                            ) : (
                              <>
                                {(project.team || []).map((initials, i) => (
                                  <span
                                    key={`${initials}-${i}`}
                                    className={css.teamChip}
                                    style={{
                                      background: AVATAR_COLORS[i % AVATAR_COLORS.length],
                                      color: AVATAR_TEXT[i % AVATAR_TEXT.length],
                                      zIndex: (project.team || []).length - i
                                    }}
                                  >
                                    {initials}
                                  </span>
                                ))}
                                {project.teamExtra > 0 && (
                                  <span className={css.teamMore}>+{project.teamExtra}</span>
                                )}
                              </>
                            )}
                          </div>
                        </td>
                        <td className={css.colAction}>
                          <span className={css.rowChevron} aria-hidden="true">
                            <svg viewBox="0 0 24 24">
                              <path
                                d="M9 6l6 6-6 6"
                                stroke="currentColor"
                                strokeWidth="2"
                                strokeLinecap="round"
                                strokeLinejoin="round"
                                fill="none"
                              />
                            </svg>
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            <div className={css.cardFoot}>
              <span>
                Showing <strong>{projectsToRender.length}</strong> of{' '}
                <strong>{projectsToRender.length}</strong> projects
              </span>
            </div>
          </div>
        </section>
      </main>
    </div>
  );
};

export default Dashboard;
