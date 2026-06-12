import * as React from 'react';
import styles from './navbar.module.scss';
import { fetchPendingApprovalsCount } from '../Approvals/approvalData';
import { useSidebar } from './SidebarContext';

type NavKey = 'dashboard' | 'employee' | 'approvals';

export interface INavBarProps {
  activeItem?: NavKey;
  onNavigate?: (view: string) => void;
  pendingApprovalsCount?: number;
  refreshKey?: number;
}

interface NavItem {
  key: NavKey;
  label: string;
  icon: JSX.Element;
}

const DashboardIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <rect x="4" y="4" width="7" height="7" rx="1.5" />
    <rect x="13" y="4" width="7" height="7" rx="1.5" />
    <rect x="4" y="13" width="7" height="7" rx="1.5" />
    <rect x="13" y="13" width="7" height="7" rx="1.5" />
  </svg>
);

const PeopleIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <circle cx="9" cy="8" r="3" />
    <path d="M3.5 18c0-2.6 2.4-4.5 5.5-4.5s5.5 1.9 5.5 4.5" />
    <circle cx="17" cy="9" r="2.5" />
    <path d="M14 17.5c.7-1.6 2.3-2.7 4.5-2.7 1.3 0 2.3.4 3 .9" />
  </svg>
);

const ApprovalIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2" />
    <rect x="9" y="3" width="6" height="4" rx="1" />
    <path d="M9 12l2 2 4-4" />
  </svg>
);

const CollapseIcon: React.FC<{ collapsed: boolean }> = ({ collapsed }) => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    {collapsed ? (
      <path d="M10 7l5 5-5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" fill="none" />
    ) : (
      <path d="M14 17l-5-5 5-5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" fill="none" />
    )}
  </svg>
);

const navItems: NavItem[] = [
  { key: 'dashboard',  label: 'Dashboard',  icon: <DashboardIcon />  },
  { key: 'employee',   label: 'Employee',   icon: <PeopleIcon />     },
  { key: 'approvals',  label: 'Approvals',  icon: <ApprovalIcon />   }
];

const NavBar: React.FC<INavBarProps> = ({
  activeItem = 'dashboard',
  onNavigate,
  pendingApprovalsCount,
  refreshKey = 0
}) => {
  const { collapsed, toggleCollapsed } = useSidebar();
  const [pendingCount, setPendingCount] = React.useState(0);

  React.useEffect(() => {
    if (pendingApprovalsCount !== undefined) {
      setPendingCount(pendingApprovalsCount);
      return;
    }

    let cancelled = false;
    fetchPendingApprovalsCount()
      .then((count) => {
        if (!cancelled) setPendingCount(count);
      })
      .catch(() => {
        if (!cancelled) setPendingCount(0);
      });

    return () => { cancelled = true; };
  }, [pendingApprovalsCount, activeItem, refreshKey]);

  return (
    <aside
      className={`${styles.sidebar} ${collapsed ? styles.sidebarCollapsed : ''}`}
      aria-label="Main navigation"
    >
      <div className={styles.brand}>
        <div className={styles.logoWrap}>
          <svg viewBox="0 0 24 24" aria-hidden="true">
            <path d="M12 3l8 4.5v9L12 21l-8-4.5v-9L12 3z" />
            <path d="M8.5 9.5h7M8.5 12h7M8.5 14.5h4.5" />
          </svg>
        </div>
        <div className={styles.brandText}>
          <h1 className={styles.brandTitle}>AllocationIQ</h1>
        </div>
      </div>

      <nav className={styles.menu}>
        {navItems.map((item) => {
          const active = item.key === activeItem;
          return (
            <button
              key={item.key}
              type="button"
              className={`${styles.menuItem} ${active ? styles.active : ''} ${item.key === 'approvals' ? styles.approvalsItem : ''}`}
              onClick={() => onNavigate && onNavigate(item.key)}
              title={collapsed ? item.label : undefined}
              aria-label={item.label}
            >
              <span className={styles.menuIcon}>{item.icon}</span>
              <span className={styles.menuLabel}>{item.label}</span>
              {item.key === 'approvals' && pendingCount > 0 && (
                <span
                  className={styles.approvalsBadge}
                  aria-label={`${pendingCount} pending approvals`}
                >
                  {collapsed ? '' : pendingCount}
                </span>
              )}
            </button>
          );
        })}
      </nav>

      <button
        type="button"
        className={styles.collapseBtn}
        onClick={toggleCollapsed}
        aria-expanded={!collapsed}
        aria-label={collapsed ? 'Expand sidebar' : 'Collapse sidebar'}
        title={collapsed ? 'Expand sidebar' : 'Collapse sidebar'}
      >
        <CollapseIcon collapsed={collapsed} />
        <span className={styles.collapseLabel}>{collapsed ? 'Expand' : 'Collapse'}</span>
      </button>
    </aside>
  );
};

export default NavBar;
