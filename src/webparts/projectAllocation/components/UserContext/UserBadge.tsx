import * as React from 'react';
import { useUser } from './UserContext';

export interface IUserBadgeClassNames {
  userBadge: string;
  userText: string;
  userName: string;
  userMail: string;
  userAvatar: string;
}

export interface IUserBadgeProps {
  classNames: IUserBadgeClassNames;
}

const UserBadge: React.FC<IUserBadgeProps> = ({ classNames: css }) => {
  const { displayName, email, initials, greeting } = useUser();

  return (
    <div className={css.userBadge}>
      <div className={css.userText}>
        <p className={css.userName}>{greeting}, {displayName}</p>
        {email ? <p className={css.userMail}>{email}</p> : null}
      </div>
      <div className={css.userAvatar} aria-hidden="true">{initials}</div>
    </div>
  );
};

export default UserBadge;
