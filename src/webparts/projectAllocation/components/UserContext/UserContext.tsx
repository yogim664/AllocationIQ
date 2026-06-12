import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUserContextValue {
  displayName: string;
  email: string;
  initials: string;
  greeting: string;
}

const defaultValue: IUserContextValue = {
  displayName: 'User',
  email: '',
  initials: 'U',
  greeting: 'Hi'
};

const UserContext = React.createContext<IUserContextValue>(defaultValue);

const getInitials = (fullName: string): string => {
  if (!fullName) return 'U';
  const parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
};

export interface IUserContextProviderProps {
  context: WebPartContext;
  children: React.ReactNode;
}

export const UserContextProvider: React.FC<IUserContextProviderProps> = ({ context, children }) => {
  const displayName = context.pageContext.user.displayName || 'User';
  const email = context.pageContext.user.email || '';

  const value = React.useMemo<IUserContextValue>(() => ({
    displayName,
    email,
    initials: getInitials(displayName),
    greeting: 'Hi'
  }), [displayName, email]);

  return (
    <UserContext.Provider value={value}>
      {children}
    </UserContext.Provider>
  );
};

export const useUser = (): IUserContextValue => React.useContext(UserContext);
