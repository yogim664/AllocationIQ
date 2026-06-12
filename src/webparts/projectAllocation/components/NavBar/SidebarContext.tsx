import * as React from 'react';

const STORAGE_KEY = 'aegis-sidebar-collapsed';
const SIDEBAR_WIDTH_EXPANDED = '230px';
const SIDEBAR_WIDTH_COLLAPSED = '72px';

interface ISidebarContextValue {
  collapsed: boolean;
  toggleCollapsed: () => void;
}

const SidebarContext = React.createContext<ISidebarContextValue>({
  collapsed: false,
  toggleCollapsed: () => { /* noop */ }
});

const readCollapsed = (): boolean => {
  try {
    return localStorage.getItem(STORAGE_KEY) === 'true';
  } catch {
    return false;
  }
};

const applySidebarWidth = (collapsed: boolean): void => {
  const width = collapsed ? SIDEBAR_WIDTH_COLLAPSED : SIDEBAR_WIDTH_EXPANDED;
  document.documentElement.style.setProperty('--aegis-sidebar-width', width);
  document.documentElement.setAttribute('data-sidebar-collapsed', collapsed ? 'true' : 'false');
};

export const SidebarProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [collapsed, setCollapsed] = React.useState(() => {
    const initial = readCollapsed();
    if (typeof document !== 'undefined') {
      applySidebarWidth(initial);
    }
    return initial;
  });

  const toggleCollapsed = React.useCallback((): void => {
    setCollapsed((prev) => {
      const next = !prev;
      try {
        localStorage.setItem(STORAGE_KEY, String(next));
      } catch {
       
      }
      return next;
    });
  }, []);

  React.useEffect(() => {
    applySidebarWidth(collapsed);
  }, [collapsed]);

  const value = React.useMemo(
    () => ({ collapsed, toggleCollapsed }),
    [collapsed, toggleCollapsed]
  );

  return (
    <SidebarContext.Provider value={value}>
      {children}
    </SidebarContext.Provider>
  );
};

export const useSidebar = (): ISidebarContextValue => React.useContext(SidebarContext);
