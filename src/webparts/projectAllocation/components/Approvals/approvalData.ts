import { getSP } from '../../../../service/initservice';

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

/** Count of projects in the Pending Approval Queue (isManagerApproved === false). */
export async function fetchPendingApprovalsCount(): Promise<number> {
  const sp = getSP();

  try {
    const allItems = await sp.web.lists
      .getByTitle('Project')
      .items
      .select('Id', 'isManagerApproved')
      .top(5000)();
    return allItems.filter(isManagerApprovedFalse).length;
  } catch {
    return 0;
  }
}
