import * as React from 'react';
import type { IProjectAllocationProps } from './IProjectAllocationProps';
import ProjectIntakeForm from './Form/form';
import Dashboard, { ProjectRow } from './Dashboard/dashboard';
import ProjectDetails from './ProjectDetails/ProjectDetails';
import Approvals from './Approvals/approvals';
import Employee from './Employee/employee';
import { UserContextProvider } from './UserContext/UserContext';
import { SidebarProvider } from './NavBar/SidebarContext';
import '../assets/style.css';

type AppView = 'dashboard' | 'form' | 'details' | 'approvals' | 'employee';

const ProjectAllocation: React.FC<IProjectAllocationProps> = ({ context }) => {
  const [view, setView] = React.useState<AppView>('dashboard');
  const [selectedProject, setSelectedProject] = React.useState<ProjectRow | null>(null);
  const [dashboardRefreshKey, setDashboardRefreshKey] = React.useState(0);
  const [approvalsRefreshKey, setApprovalsRefreshKey] = React.useState(0);

  const handleFormSuccess = (): void => {
    setView('dashboard');
    setDashboardRefreshKey((key) => key + 1);
    setApprovalsRefreshKey((key) => key + 1);
  };

  const handleNavigate = (target: string): void => {
    if (target === 'dashboard' || target === 'approvals' || target === 'employee') {
      setView(target as AppView);
    }
  };

  return (
    <UserContextProvider context={context}>
    <SidebarProvider>
    <div
      className="aegisApp"
      style={{
      position: 'relative',
      minHeight: '100vh',
      padding: 0,
      margin: 0,
      width: '100%',
      maxWidth: '100%',
      display: 'flex',
      flexDirection: 'column',
      boxSizing: 'border-box',
      overflow: 'visible'
    }}>

      {view === 'details' && selectedProject ? (
        <ProjectDetails project={selectedProject} onBack={() => setView('dashboard')} />
      ) : view === 'approvals' ? (
        <Approvals
          onNavigate={handleNavigate}
          onApprovalSuccess={() => {
            setView('dashboard');
            setDashboardRefreshKey((key) => key + 1);
            setApprovalsRefreshKey((key) => key + 1);
          }}
        />
      ) : view === 'employee' ? (
        <Employee onNavigate={handleNavigate} approvalsRefreshKey={approvalsRefreshKey} />
      ) : (
        <Dashboard
          onCreateProject={() => setView('form')}
          onOpenProject={(project) => {
            setSelectedProject(project);
            setView('details');
          }}
          onNavigate={handleNavigate}
          refreshKey={dashboardRefreshKey}
          approvalsRefreshKey={approvalsRefreshKey}
        />
      )}

      {view === 'form' && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(15, 23, 42, 0.6)',
          backdropFilter: 'blur(8px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999,
          padding: '24px 16px'
        }}>
          <div style={{
            width: '90vw',
            maxWidth: '1480px',
            height: '90vh',
            maxHeight: '860px',
            position: 'relative',
            display: 'flex',
            flexDirection: 'column'
          }}>
            <ProjectIntakeForm
              onClose={() => setView('dashboard')}
              onSuccess={handleFormSuccess}
            />
          </div>
        </div>
      )}
    </div>
    </SidebarProvider>
    </UserContextProvider>
  );
};

export default ProjectAllocation;
