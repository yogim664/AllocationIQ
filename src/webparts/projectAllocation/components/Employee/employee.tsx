import * as React from 'react';
import styles from './employee.module.scss';
import NavBar from '../NavBar/navbar';
import { getSP } from '../../../../service/initservice';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IEmployeeItem {
  Id: number;
  Title: string;           // Employee ID
  field_1: string;         // FullName
  field_2: string;         // Department
  field_3: string;         // Designation
  field_4: string;         // Email
  field_5: number | null;  // Phone
  field_6: string;         // Location
  field_7: string;         // JoiningDate
  field_8: string;         // EmploymentType
  field_9: string;         // ManagerID
  field_10: string;        // Status
}

export interface IEmployeeProps {
  onNavigate?: (view: string) => void;
  approvalsRefreshKey?: number;
}

const getInitials = (name: string): string => {
  if (!name) return '??';
  const parts = name.trim().split(/\s+/);
  if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  return name.slice(0, 2).toUpperCase();
};

const getStatusClass = (status: string): string => {
  const s = (status || '').toLowerCase().trim();
  if (s === 'active') return styles.statusActive;
  if (s === 'inactive') return styles.statusInactive;
  if (s === 'on leave' || s === 'onleave') return styles.statusOnLeave;
  return styles.statusActive;
};

const PeopleIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <circle cx="9" cy="8" r="3" stroke="currentColor" strokeWidth="1.8" />
    <path d="M3.5 18c0-2.6 2.4-4.5 5.5-4.5s5.5 1.9 5.5 4.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    <circle cx="17" cy="9" r="2.5" stroke="currentColor" strokeWidth="1.8" />
    <path d="M14 17.5c.7-1.6 2.3-2.7 4.5-2.7 1.3 0 2.3.4 3 .9" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const ActiveIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <circle cx="12" cy="8" r="3.5" stroke="currentColor" strokeWidth="1.8" />
    <path d="M6 19c0-2.8 2.7-5 6-5s6 2.2 6 5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    <path d="M16 10l1.5 1.5L20 9" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const BuildingIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <rect x="5" y="4" width="14" height="16" rx="1.5" stroke="currentColor" strokeWidth="1.8" />
    <path d="M9 8h2M13 8h2M9 12h2M13 12h2M9 16h2M13 16h2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const LocationIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <path d="M12 21s6-5.2 6-10a6 6 0 10-12 0c0 4.8 6 10 6 10z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
    <circle cx="12" cy="11" r="2.5" stroke="currentColor" strokeWidth="1.8" />
  </svg>
);

const Employee: React.FC<IEmployeeProps> = ({ onNavigate, approvalsRefreshKey = 0 }) => {
  const css = styles as typeof styles & Record<string, string>;

  const [employees, setEmployees] = React.useState<IEmployeeItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');
  const [searchTerm, setSearchTerm] = React.useState<string>('');

  const fetchEmployees = React.useCallback(async (): Promise<void> => {
    setLoading(true);
    setError('');
    try {
      const sp = getSP();
      const items: IEmployeeItem[] = await sp.web.lists
        .getByTitle("Employees")
        .items
        .select("*")
        .top(5000)();
      setEmployees(items);
    } catch (err) {
      console.error("Error fetching employees:", err);
      setError("Failed to load employee data. Please try again.");
    } finally {
      setLoading(false);
    }
  }, []);

  React.useEffect(() => {
    fetchEmployees().catch(console.error);
  }, [fetchEmployees]);

  const filtered = React.useMemo(() => {
    if (!searchTerm.trim()) return employees;
    const q = searchTerm.toLowerCase();
    return employees.filter(
      (emp) =>
        (emp.field_1 || '').toLowerCase().includes(q) ||
        (emp.Title || '').toLowerCase().includes(q) ||
        (emp.field_2 || '').toLowerCase().includes(q) ||
        (emp.field_3 || '').toLowerCase().includes(q) ||
        (emp.field_4 || '').toLowerCase().includes(q) ||
        (emp.field_6 || '').toLowerCase().includes(q) ||
        (emp.field_10 || '').toLowerCase().includes(q)
    );
  }, [employees, searchTerm]);

  const stats = React.useMemo(() => {
    const total = employees.length;
    const active = employees.filter((e) => (e.field_10 || '').toLowerCase() === 'active').length;
    const departments = new Set(employees.map((e) => e.field_2).filter(Boolean)).size;
    const locations = new Set(employees.map((e) => e.field_6).filter(Boolean)).size;
    return { total, active, departments, locations };
  }, [employees]);

  return (
    <div className={css.page}>
      <div className={css.pageMesh} aria-hidden="true" />
      <div className={css.pageOrb} aria-hidden="true" />
      <div className={css.pageOrb2} aria-hidden="true" />

      <NavBar activeItem="employee" onNavigate={onNavigate} refreshKey={approvalsRefreshKey} />

      <main className={css.main}>
        <section className={css.statsRow} aria-label="Employee summary">
          <article className={`${css.statCard} ${css.statTotal}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><PeopleIcon /></span>
            </div>
            <p className={css.statLabel}>Total Employees</p>
            <p className={css.statValue}>{stats.total}</p>
            <span className={`${css.statTrend} ${css.statTrendPurple}`}>All Records</span>
            <div className={`${css.statWatermark} ${css.statWatermarkPurple}`} aria-hidden="true">
              <PeopleIcon />
            </div>
          </article>
          <article className={`${css.statCard} ${css.statActive}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><ActiveIcon /></span>
            </div>
            <p className={css.statLabel}>Active Employees</p>
            <p className={css.statValue}>{stats.active}</p>
            <span className={`${css.statTrend} ${css.statTrendGreen}`}>Currently Active</span>
            <div className={`${css.statWatermark} ${css.statWatermarkGreen}`} aria-hidden="true">
              <ActiveIcon />
            </div>
          </article>
          <article className={`${css.statCard} ${css.statDept}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><BuildingIcon /></span>
            </div>
            <p className={css.statLabel}>Departments</p>
            <p className={css.statValue}>{stats.departments}</p>
            <span className={`${css.statTrend} ${css.statTrendBlue}`}>Across Organization</span>
            <div className={`${css.statWatermark} ${css.statWatermarkBlue}`} aria-hidden="true">
              <BuildingIcon />
            </div>
          </article>
          <article className={`${css.statCard} ${css.statLocation}`}>
            <div className={css.statCardBar} aria-hidden="true" />
            <div className={css.statCardHead}>
              <span className={css.statIcon} aria-hidden="true"><LocationIcon /></span>
            </div>
            <p className={css.statLabel}>Locations</p>
            <p className={css.statValue}>{stats.locations}</p>
            <span className={`${css.statTrend} ${css.statTrendAmber}`}>Office Sites</span>
            <div className={`${css.statWatermark} ${css.statWatermarkAmber}`} aria-hidden="true">
              <LocationIcon />
            </div>
          </article>
        </section>

        <section className={css.directoryCard}>
          <header className={css.directoryHead}>
            <div className={css.directoryHeadText}>
              <h1 className={css.directoryTitle}>Employee Directory</h1>
              <p className={css.directorySubtitle}>
                Browse and manage employee records across the organization.
              </p>
            </div>
            <div className={css.searchBox}>
              <svg className={css.searchIcon} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <circle cx="11" cy="11" r="8" />
                <path d="M21 21l-4.35-4.35" />
              </svg>
              <input
                type="text"
                className={css.searchInput}
                placeholder="Search employees..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                aria-label="Search employees"
              />
            </div>
          </header>

          {loading && (
            <div className={css.loadingWrap}>
              <div className={css.spinner} />
              <span className={css.loadingText}>Loading employee data…</span>
            </div>
          )}

          {!loading && error && (
            <div className={css.errorText}>{error}</div>
          )}

          {!loading && !error && filtered.length === 0 && (
            <div className={css.emptyWrap}>
              <svg className={css.emptyIcon} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                <circle cx="9" cy="8" r="3" />
                <path d="M3.5 18c0-2.6 2.4-4.5 5.5-4.5s5.5 1.9 5.5 4.5" />
                <circle cx="17" cy="9" r="2.5" />
                <path d="M14 17.5c.7-1.6 2.3-2.7 4.5-2.7 1.3 0 2.3.4 3 .9" />
              </svg>
              <span className={css.emptyText}>
                {searchTerm ? 'No employees match your search' : 'No employee records found'}
              </span>
              <span className={css.emptySub}>
                {searchTerm ? 'Try a different search term' : 'Employee data will appear here once added'}
              </span>
            </div>
          )}

          {!loading && !error && filtered.length > 0 && (
            <div className={css.tableWrap}>
              <table className={css.table}>
                <thead>
                  <tr>
                    <th>Employee</th>
                    <th>Department</th>
                    <th>Designation</th>
                    <th>Email</th>
                    <th>Location</th>
                    <th>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((emp) => (
                    <tr key={emp.Id} className={css.empRow} tabIndex={0}>
                      <td className={css.nameCell}>
                        <div className={css.nameRow}>
                          <span className={css.avatar}>
                            {getInitials(emp.field_1 || emp.Title)}
                          </span>
                          <div className={css.nameText}>
                            <strong>{emp.field_1 || '—'}</strong>
                            <span>{emp.field_3 || '—'}</span>
                          </div>
                        </div>
                      </td>
                      <td>{emp.field_2 || '—'}</td>
                      <td>{emp.field_3 || '—'}</td>
                      <td className={css.emailCell}>{emp.field_4 || '—'}</td>
                      <td>{emp.field_6 || '—'}</td>
                      <td>
                        <span className={`${css.statusPill} ${getStatusClass(emp.field_10)}`}>
                          {emp.field_10 || 'Active'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>
      </main>
    </div>
  );
};

export default Employee;
