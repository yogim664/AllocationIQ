import * as React from 'react';
import styles from './form.module.scss';
import * as XLSX from 'xlsx';
import { uploadPdfAndAnalyze, StaffRecommandAgent, ProjectPlanAgent } from '../AI/aiagents';
import { getSP } from '../../../../service/initservice';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const STEPS = [
  'SOURCE',
  'INTAKE SPEC',
  'STAFF RECOMMENDATION',
  'PROJECT PLAN',
  'APPROVALS',
  'SUCCESS NOTES'
] as const;

const SKILL_OPTIONS = [
  'React',
  'Azure',
  'Python',
  'AI/ML',
  "Azure AI Foundry",
  "Microsoft Graph API",
  "PnPjs",
  "Power Automate",
  "Microsoft Teams",
  "SharePoint",
  "Copilot Studio",
  "Power BI",
  'Node.js',
  'TypeScript',
  'DevOps',
  'Power Platform'
];

const PRIORITIES = ['Low', 'Medium', 'High', 'Critical'] as const;
type Priority = (typeof PRIORITIES)[number];

interface StaffMember {
  EmployeeID:string,
  initials: string;
  name: string;
  role: string;
  department: string;
  match: number;
  availability: number;
  success: number;
}

type PlanLane = 'development' | 'demo' | 'testing';
type PlanStatus = 'Not Started' | 'In Progress' | 'Done';

interface PlanItem {
  slNo: number;
  feature: string;
  task: string;
  description: string;
  startDate: string;
  endDate: string;
  efforts: number;
  lane: PlanLane;
  startSlot: number;
  durationSlots: number;
  assignee: string;
  status: PlanStatus;
}

const PLAN_STATUSES: PlanStatus[] = ['Not Started', 'In Progress', 'Done'];

const getStaffInitials = (name: string): string => {
  if (!name) return '??';
  return name
    .split(' ')
    .map((p) => p[0])
    .join('')
    .slice(0, 2)
    .toUpperCase();
};

const normalizeStaffMember = (item: Record<string, unknown>, index: number): StaffMember => ({
  EmployeeID: String(item.EmployeeID ?? item.employeeId ?? item.id ?? `Emp${index + 1}`),
  initials: String(item.initials ?? getStaffInitials(String(item.name ?? item.Name ?? ''))),
  name: String(item.name ?? item.Name ?? item.fullName ?? ''),
  role: String(item.role ?? item.Role ?? item.designation ?? ''),
  department: String(item.department ?? item.Department ?? ''),
  match: Number(item.match ?? item.Match ?? item.matchScore ?? 0),
  availability: Number(item.availability ?? item.Availability ?? 0),
  success: Number(item.success ?? item.Success ?? item.successRate ?? 0),
});

const parseAgentJson = (data: { response?: string | Record<string, unknown> } | null | undefined): unknown => {
  if (!data) return null;
  if (typeof data.response === 'object' && data.response !== null) return data.response;
  if (typeof data.response !== 'string') return data;
  try {
    const jsonMatch = data.response.match(/```(?:json)?\s*([\s\S]*?)```/);
    const jsonText = (jsonMatch ? jsonMatch[1] : data.response).trim();
    return JSON.parse(jsonText);
  } catch {
    return null;
  }
};

const getAssigneeInitials = (name: string, pool: StaffMember[]): string => {
  const found = pool.find((s) => s.name === name);
  if (found) return found.initials;
  return name
    .split(' ')
    .map((p) => p[0])
    .join('')
    .slice(0, 2)
    .toUpperCase();
};

const parsePlanDate = (dateStr: string): Date => {
  if (!dateStr) return new Date(0);
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const monthStr = parts[1];
    let year = parseInt(parts[2], 10);
    if (year < 100) year += 2000;
    const months: Record<string, number> = {
      jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
      jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
    };
    const month = months[monthStr.toLowerCase()];
    if (month !== undefined && !isNaN(day) && !isNaN(year)) {
      return new Date(year, month, day);
    }
  }
  return new Date(0);
};

const AI_LOADING_MS = 10000;

interface IntakeState {
  projectName: string;
  clientName: string;
  budget: string;
  startDate: string;
  endDate: string;
  projectDescription: string;
  priority: Priority;
  skills: string[];
  architectNotes: string;
}

const defaultIntake: IntakeState = {
  projectName: '',
  clientName: '',
  budget: '',
  startDate: '',
  endDate: '',
  projectDescription: '',
  priority: 'Low',
  skills: [],
  architectNotes: ''
};

const ChipIcon: React.FC = () => (
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <rect x="4" y="4" width="16" height="16" rx="3" stroke="white" strokeWidth="1.5" />
    <path d="M9 9h6v6H9z" fill="white" opacity="0.9" />
    <path d="M12 4v3M12 17v3M4 12h3M17 12h3" stroke="white" strokeWidth="1.2" strokeLinecap="round" />
  </svg>
);

const CheckIcon: React.FC = () => (
  <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
    <path d="M3 7l3 3 5-6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const SourceDiamondIcon: React.FC = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M12 3l7 4v10l-7 4-7-4V7l7-4z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
  </svg>
);

const DocumentUploadIcon: React.FC = () => (
  <svg width="26" height="26" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M8 4h8l2 4v12a1 1 0 01-1 1H7a1 1 0 01-1-1V4z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
    <path d="M8 4v4h10M12 12v5M9.5 14.5L12 17l2.5-2.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const ManualEntryIcon: React.FC = () => (
  <svg width="26" height="26" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M4 20h4l10-10-4-4L4 16v4z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
    <path d="M13.5 6.5l2 2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const CardArrowIcon: React.FC = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path d="M5 12h14M13 6l6 6-6 6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const ExternalLinkIcon: React.FC = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
    <path d="M6 3h7v7M13 3L6 10M9 3H3v10h10V9" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const SuccessCheckIcon: React.FC = () => (
  <svg className={styles.successCheckSvg} width="36" height="36" viewBox="0 0 36 36" fill="none" aria-hidden="true">
    <path
      className={styles.successCheckPath}
      d="M8 18l7 7 13-14"
      stroke="white"
      strokeWidth="3"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

const PlanEffortIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.8" />
    <path d="M12 7v5l3 2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const PlanDurationIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <rect x="4" y="5" width="16" height="15" rx="2" stroke="currentColor" strokeWidth="1.8" />
    <path d="M8 3v4M16 3v4M4 10h16" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

const PlanTeamIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true" fill="none">
    <circle cx="9" cy="8" r="3" stroke="currentColor" strokeWidth="1.8" />
    <path d="M3.5 18c0-2.6 2.4-4.5 5.5-4.5s5.5 1.9 5.5 4.5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
    <circle cx="17" cy="9" r="2.5" stroke="currentColor" strokeWidth="1.8" />
    <path d="M14 17.5c.7-1.6 2.3-2.7 4.5-2.7 1.3 0 2.3.4 3 .9" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
  </svg>
);

interface AgentDef {
  id: string;
  label: string;
  task: string;
}

const AI_AGENTS: AgentDef[] = [

];

const PLAN_GENERATION_AGENTS: AgentDef[] = [

];

const PLAN_LOADING_TITLE = 'Generate an AI Powered Project Plan';
const PLAN_LOADING_MESSAGE = 'AI is building your project plan and timeline phases…';

const SUBMIT_AGENTS: AgentDef[] = [

];

const SUBMIT_LOADING_TITLE = 'Submitting for Approval';
const SUBMIT_LOADING_MESSAGE = 'Provisioning project records and submitting for manager review…';

const DOC_ANALYSIS_AGENTS: AgentDef[] = [
  
];
const AI_AGENT_PHASES = AI_AGENTS.map((a) => a.task);
const DOC_ANALYSIS_MS = 6000;
const SOURCE_ADVANCE_MS = 2800;

type SourceMode = 'upload' | 'manual';
type SourcePhase = 'choose' | 'upload' | 'analyzing' | 'complete';
const formatDisplayDate = (dateStr: string): string => {
  if (!dateStr) return '—';
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  } catch {
    return dateStr;
  }
};

const parseIntakeDate = (dateStr: string): Date => {
  if (!dateStr) return new Date(NaN);
  const parts = dateStr.split('-');
  if (parts.length === 3 && parts[0].length === 4) {
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    if (!isNaN(year) && !isNaN(month) && !isNaN(day)) {
      return new Date(year, month, day);
    }
  }
  const planDate = parsePlanDate(dateStr);
  return planDate.getTime() === 0 ? new Date(NaN) : planDate;
};

const getDurationDays = (start: string, end: string): number => {
  const s = parseIntakeDate(start);
  const e = parseIntakeDate(end);
  if (isNaN(s.getTime()) || isNaN(e.getTime())) return 0;
  const days = Math.round(Math.abs(e.getTime() - s.getTime()) / (1000 * 60 * 60 * 24));
  return Math.max(1, days);
};

const getDurationMonths = (start: string, end: string): string => {
  const s = parseIntakeDate(start);
  const e = parseIntakeDate(end);
  if (isNaN(s.getTime()) || isNaN(e.getTime())) return '—';
  const months = Math.max(
    1,
    Math.round(Math.abs(e.getTime() - s.getTime()) / (1000 * 60 * 60 * 24 * 30.4))
  );
  return `${months} Month${months !== 1 ? 's' : ''}`;
};

const AiSparkleLoader: React.FC = () => (
  <div className={styles.aiSparkleLoader} aria-hidden="true" role="status">
   
    <svg className={styles.aiOrbitRing1} viewBox="0 0 120 120" fill="none">
      <circle cx="60" cy="60" r="54" stroke="rgba(124,58,237,0.18)" strokeWidth="1.5" strokeDasharray="8 6" />
    </svg>
   
    <svg className={styles.aiOrbitRing2} viewBox="0 0 120 120" fill="none">
      <circle cx="60" cy="60" r="42" stroke="rgba(37,99,235,0.2)" strokeWidth="1.2" strokeDasharray="5 8" />
    </svg>
  
    <div className={styles.aiOrbCore}>
      <div className={styles.aiOrbGlow} />
      <svg className={styles.aiSparkleSvg} viewBox="0 0 60 60" fill="none">
        <path
          className={styles.aiSparkleMain}
          d="M30 16l2.8 11.2L44 30l-11.2 2.8L30 44l-2.8-11.2L16 30l11.2-2.8L30 16z"
          fill="#7c3aed"
        />
        <path
          className={styles.aiSparkleDot1}
          d="M46 14l1.4 2.8 2.8 1.4-2.8 1.4L46 22l-1.4-2.8L41.8 18l2.8-1.4L46 14z"
          fill="#a78bfa"
        />
        <path
          className={styles.aiSparkleDot2}
          d="M16 40l1 2 2 1-2 1-1 2-1-2-2-1 2-1 1-2z"
          fill="#818cf8"
        />
      </svg>
    </div>
  
    <div className={styles.aiOrbitDot1} />
 
 
    <div className={styles.aiOrbitDot2} />
  </div>
);

const getWeekday = (dateStr: string): string => {
  if (!dateStr) return '';
  const parts = dateStr.split('-');
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const monthStr = parts[1];
    let year = parseInt(parts[2], 10);
    if (year < 100) year += 2000;
    const months: Record<string, number> = {
      jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
      jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
    };
    const month = months[monthStr.toLowerCase()];
    if (month !== undefined && !isNaN(day) && !isNaN(year)) {
      const date = new Date(year, month, day);
      const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
      return days[date.getDay()];
    }
  }
  return '';
};

const getExcelColumnLabel = (index: number): string => {
  let label = '';
  let temp = index;
  while (temp >= 0) {
    label = String.fromCharCode((temp % 26) + 65) + label;
    temp = Math.floor(temp / 26) - 1;
  }
  return label;
};

const getRowSpans = (items: PlanItem[]) => {
  const slSpans: number[] = [];
  const featSpans: number[] = [];
  
  let i = 0;
  while (i < items.length) {
    let j = i + 1;
    while (j < items.length && items[j].slNo === items[i].slNo) {
      j++;
    }
    const span = j - i;
    slSpans[i] = span;
    for (let k = i + 1; k < j; k++) {
      slSpans[k] = 0;
    }
    i = j;
  }
  
  i = 0;
  while (i < items.length) {
    let j = i + 1;
    while (j < items.length && items[j].feature === items[i].feature) {
      j++;
    }
    const span = j - i;
    featSpans[i] = span;
    for (let k = i + 1; k < j; k++) {
      featSpans[k] = 0;
    }
    i = j;
  }
  
  return { slSpans, featSpans };
};

interface MonthSpan {
  label: string;
  span: number;
}

const getMonthSpans = (columns: string[]): MonthSpan[] => {
  const spans: MonthSpan[] = [];
  if (columns.length === 0) return spans;
  
  let currentMonthStr = '';
  let currentSpan = 0;
  
  columns.forEach((col) => {
    const parts = col.split('-');
    let monthLabel = '';
    if (parts.length === 3) {
      const m = parts[1];
      let y = parseInt(parts[2], 10);
      if (y < 100) y += 2000;
      const monthNames: Record<string, string> = {
        jan: 'Jan', feb: 'Feb', mar: 'Mar', apr: 'Apr',
        may: 'May', jun: 'Jun', jul: 'Jul', aug: 'Aug',
        sep: 'Sep', oct: 'Oct', nov: 'Nov', dec: 'Dec'
      };
      const monthShort = monthNames[m.toLowerCase()] || m;
      monthLabel = `${monthShort} ${y}`;
    } else {
      monthLabel = col;
    }
    
    if (monthLabel === currentMonthStr) {
      currentSpan++;
    } else {
      if (currentSpan > 0) {
        spans.push({ label: currentMonthStr, span: currentSpan });
      }
      currentMonthStr = monthLabel;
      currentSpan = 1;
    }
  });
  
  if (currentSpan > 0) {
    spans.push({ label: currentMonthStr, span: currentSpan });
  }
  
  return spans;
};

interface IProjectIntakeFormProps {
  onClose?: () => void;
  onSuccess?: () => void;
  onStepChange?: (step: number) => void;
}

const SUCCESS_REDIRECT_MS = 2500;

const ProjectIntakeForm: React.FC<IProjectIntakeFormProps> = ({ onClose, onSuccess, onStepChange }) => {
  const css = styles as typeof styles & Record<string, string>;
  const [step, setStep] = React.useState(1);
  const [intake, setIntake] = React.useState<IntakeState>(defaultIntake);
  const [sourceMode, setSourceMode] = React.useState<SourceMode | null>(null);
  const [sourcePhase, setSourcePhase] = React.useState<SourcePhase>('choose');
  const [uploadedFileName, setUploadedFileName] = React.useState<string | null>(null);
  const [aiLoading, setAiLoading] = React.useState(false);
  const [aiProgress, setAiProgress] = React.useState(0);
  const [aiMessage, setAiMessage] = React.useState('');
  const [recommendedStaff, setRecommendedStaff] = React.useState<StaffMember[]>([]);
  const [planItems, setPlanItems] = React.useState<PlanItem[]>([]);
  const [timelineColumns, setTimelineColumns] = React.useState<string[]>([]);
  const [expandedFeatures, setExpandedFeatures] = React.useState<Record<string, boolean>>({});
  const [planZoom, setPlanZoom] = React.useState(100);
  const sourceFileRef = React.useRef<HTMLInputElement>(null);
  const sourceAdvanceRef = React.useRef<number | undefined>();

  React.useEffect(() => {
    onStepChange?.(step);
  }, [step, onStepChange]);

  React.useEffect(() => {
    if (step !== 6) return undefined;
    const complete = onSuccess ?? onClose;
    if (!complete) return undefined;

    const timer = window.setTimeout(() => {
      complete();
    }, SUCCESS_REDIRECT_MS);

    return () => { window.clearTimeout(timer); };
  }, [step, onSuccess, onClose]);

  const toggleFeatureExpand = (feat: string) => {
    setExpandedFeatures((prev) => ({
      ...prev,
      [feat]: !prev[feat]
    }));
  };
  const loadingTimerRef = React.useRef<number | undefined>();
  const progressTimerRef = React.useRef<number | undefined>();
  const submitProgressRef = React.useRef<number | undefined>();

  const clearTimers = React.useCallback((): void => {
    if (loadingTimerRef.current) {
      window.clearTimeout(loadingTimerRef.current);
      loadingTimerRef.current = undefined;
    }
    if (progressTimerRef.current) {
      window.clearInterval(progressTimerRef.current);
      progressTimerRef.current = undefined;
    }
    if (submitProgressRef.current) {
      window.clearInterval(submitProgressRef.current);
      submitProgressRef.current = undefined;
    }
  }, []);

  React.useEffect(() => (): void => clearTimers(), [clearTimers]);

  const resetSourceStep = React.useCallback((): void => {
    setSourceMode(null);
    setSourcePhase('choose');
    setUploadedFileName(null);
    if (sourceFileRef.current) {
      sourceFileRef.current.value = '';
    }
  }, []);

  React.useEffect(() => {
    if (step !== 1 || sourcePhase !== 'complete') return undefined;
    sourceAdvanceRef.current = window.setTimeout(() => {
      setStep(2);
      resetSourceStep();
    }, SOURCE_ADVANCE_MS);
    return () => {
      if (sourceAdvanceRef.current) {
        window.clearTimeout(sourceAdvanceRef.current);
        sourceAdvanceRef.current = undefined;
      }
    };
  }, [step, sourcePhase, resetSourceStep]);

  const [loadingAgents, setLoadingAgents] = React.useState<AgentDef[]>(AI_AGENTS);

  const runAiLoading = React.useCallback(
    (
      message: string,
      onComplete: () => void,
      durationMs: number = AI_LOADING_MS,
      agents: AgentDef[] = AI_AGENTS
    ): void => {
      clearTimers();
      setLoadingAgents(agents);
      setAiLoading(true);
      setAiMessage(message);
      setAiProgress(0);

      const start = Date.now();
      progressTimerRef.current = window.setInterval(() => {
        const elapsed = Date.now() - start;
        setAiProgress(Math.min(100, (elapsed / durationMs) * 100));
      }, 100);

      loadingTimerRef.current = window.setTimeout(() => {
        clearTimers();
        setAiLoading(false);
        setAiProgress(100);
        onComplete();
      }, durationMs);
    },
    [clearTimers]
  );

  const updateIntake = (patch: Partial<IntakeState>): void => {
    setIntake((prev) => ({ ...prev, ...patch }));
  };

  const toggleSkill = (skill: string): void => {
    setIntake((prev) => ({
      ...prev,
      skills: prev.skills.indexOf(skill) >= 0
        ? prev.skills.filter((s) => s !== skill)
        : [...prev.skills, skill]
    }));
  };

  const handleContinueFromIntake = (): void => {
    if (aiLoading) return;
    setStep(3);
    StaffRecommandAgent(intake.skills, intake)
      .then((data) => {
        const parsed = parseAgentJson(data);
        const list = Array.isArray(parsed)
          ? parsed
          : (parsed as Record<string, unknown>)?.staff
            ?? (parsed as Record<string, unknown>)?.recommendedStaff
            ?? (parsed as Record<string, unknown>)?.employees;
        if (Array.isArray(list) && list.length > 0) {
          setRecommendedStaff(list.map((item, index) => normalizeStaffMember(item as Record<string, unknown>, index)));
        }
      })
      .catch((err) => {
        console.error("Staff recommendation agent failed:", err);
      });
    runAiLoading(
      'Azure AI Foundry calculating suitability arrays...',
      () => {}
    );
  };

  const handleGeneratePlan = (): void => {
    if (aiLoading) return;

    setStep(4);

    ProjectPlanAgent(recommendedStaff, intake)
      .then((data) => {
        if (!Array.isArray(data)) return;
        const datesList: string[] = [];
        data.forEach((item: any) => {
          datesList.push(String(item.startDate));
          datesList.push(String(item.endDate));
        });
        const uniqueDates = Array.from(new Set(datesList))
          .sort((a, b) => parsePlanDate(a).getTime() - parsePlanDate(b).getTime());

        const mapped = data.map((item: any) => {
          const startIdx = uniqueDates.indexOf(item.startDate);
          const endIdx = uniqueDates.indexOf(item.endDate);
          return {
            ...item,
            description: item.description ?? '',
            startSlot: startIdx >= 0 ? startIdx : 0,
            durationSlots: endIdx >= startIdx ? endIdx - startIdx + 1 : 1,
            status: item.status ?? 'Not Started'
          };
        });

        setTimelineColumns(uniqueDates);
        setPlanItems(mapped);
      })
      .catch((err) => {
        console.error("ProjectPlanAgent failed:", err);
      });

    runAiLoading(
      PLAN_LOADING_MESSAGE,
      () => {},
      4000,
      PLAN_GENERATION_AGENTS
    );
  };

  const handleSendForApproval = (): void => {
    setStep(5);
  };

  const beginSubmitLoading = (): void => {
    clearTimers();
    setLoadingAgents(SUBMIT_AGENTS);
    setAiLoading(true);
    setAiMessage(SUBMIT_LOADING_MESSAGE);
    setAiProgress(8);
    const start = Date.now();
    submitProgressRef.current = window.setInterval(() => {
      const elapsed = Date.now() - start;
      setAiProgress((prev) => Math.min(88, Math.max(prev, 8 + (elapsed / 10000) * 80)));
    }, 120);
  };

  const finishSubmitLoading = (success: boolean): void => {
    if (submitProgressRef.current) {
      window.clearInterval(submitProgressRef.current);
      submitProgressRef.current = undefined;
    }
    if (success) {
      setAiProgress(100);
      setAiLoading(false);
      setAiMessage('');
      setStep(6);
      return;
    }
    setAiLoading(false);
    setAiProgress(0);
    setAiMessage('');
  };

  const handleApprove = async (): Promise<void> => {
    if (aiLoading) return;
    beginSubmitLoading();

    try {
      const sp = getSP();

      const generateProjectCode = (): string => {
        const now = new Date();
        const mVal = now.getMonth() + 1;
        const mm = mVal < 10 ? '0' + mVal : String(mVal);
        const yyyy = now.getFullYear();
        const randomNum = Math.floor(100 + Math.random() * 900);
        return `P-${mm}-${yyyy}-${randomNum}`;
      };

      const safeISOString = (dateStr: string): string => {
        try {
          const d = new Date(dateStr);
          return isNaN(d.getTime()) ? new Date().toISOString() : d.toISOString();
        } catch {
          return new Date().toISOString();
        }
      };

      const cleanBudget = (budgetStr: string): number => {
        const cleaned = (budgetStr || '').replace(/[^0-9.-]+/g, '');
        const num = Number(cleaned);
        return isNaN(num) ? 0 : num;
      };

      const newProjectCode = generateProjectCode();

      setAiMessage(SUBMIT_AGENTS[0].task);
      setAiProgress(20);
      await sp.web.lists.getByTitle('Project').items.add({
        ProjectCode: newProjectCode,
        ProjectName: intake.projectName,
        Client: intake.clientName,
        isManagerApproved: false,
        StartDate: safeISOString(intake.startDate),
        EndDate: safeISOString(intake.endDate),
        Priority: intake.priority,
        RequiredSkills: intake.skills.join(', '),
        Budget: cleanBudget(intake.budget),
        Description: intake.projectDescription
      });

      setAiMessage(SUBMIT_AGENTS[1].task);
      setAiProgress(45);
      await Promise.all(
        recommendedStaff.map((staff) =>
          sp.web.lists.getByTitle('ProjectParticipation').items.add({
            field_1: staff.EmployeeID,
            field_2: staff.name,
            field_3: newProjectCode,
            field_4: staff.role,
            field_5: staff.availability,
            field_6: safeISOString(intake.startDate),
            field_7: safeISOString(intake.endDate)
          })
        )
      );

      setAiMessage(SUBMIT_AGENTS[2].task);
      setAiProgress(70);
      await Promise.all(
        planItems.map((task) =>
          sp.web.lists.getByTitle('TaskDetails').items.add({
            Task: task.task,
            Sno: task.slNo.toString(),
            ProjectCode: newProjectCode,
            Assignee: task.assignee,
            lane: task.lane,
            Effort: task.efforts.toString(),
            startSlot: task.startSlot.toString(),
            durationSlots: task.durationSlots.toString(),
            Status: task.status,
            StartDate: safeISOString(intake.startDate),
            EndDate: safeISOString(intake.endDate)
          })
        )
      );

      setAiMessage(SUBMIT_AGENTS[3].task);
      setAiProgress(92);
      finishSubmitLoading(true);
    } catch (err) {
      console.error('Failed to save project item in SharePoint:', err);
      finishSubmitLoading(false);
    }
  };

  const handleBack = (): void => {
    if (aiLoading) return;
    clearTimers();
    setAiLoading(false);
    setAiProgress(0);
    setAiMessage('');
    if (sourceAdvanceRef.current) {
      window.clearTimeout(sourceAdvanceRef.current);
      sourceAdvanceRef.current = undefined;
    }
    if (step === 2) {
      setStep(1);
      resetSourceStep();
      return;
    }
    if (step > 2) {
      setStep((s) => s - 1);
    }
  };

  const handleSelectManualEntry = (): void => {
    if (aiLoading) return;
    setSourceMode('manual');
    setStep(2);
  };

  const handleSelectDocumentUpload = (): void => {
    if (aiLoading) return;
    setSourceMode('upload');
    setSourcePhase('upload');
  };

  const handleSourceFileSelected = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = e.target.files?.[0];
    if (!file || aiLoading) return;
    try {
      setUploadedFileName(file.name);
      setSourcePhase('analyzing');
      runAiLoading(
        'Azure AI agent analyzing uploaded statement of work…',
        () => {
          uploadPdfAndAnalyze(file)
          .then((data: any) => {
            const parsed = (parseAgentJson(data) ?? {}) as Record<string, unknown>;

            const projectName = String(parsed.ProjectName || parsed.projectName || '');
            const clientName = String(parsed.Client || parsed.client || parsed.clientName || '');
            
            let budget = ""
            if (parsed.Budget !== undefined && parsed.Budget !== null) {
              const bNum = Number(parsed.Budget);
              if (!isNaN(bNum)) {
                budget = bNum.toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
              } else {
                budget = String(parsed.Budget);
              }
            } else if (parsed.budget) {
              budget = String(parsed.budget);
            }

            const startDate = String(parsed.StartDate || parsed.startDate || '');
            const endDate = String(parsed.EndDate || parsed.endDate || parsed.dueDate || '');

            let priority: Priority = 'Low';
            const pRaw = String(parsed.Priority || parsed.priority || 'Low').toLowerCase();
            if (pRaw === 'low') priority = 'Low';
            else if (pRaw === 'medium') priority = 'Medium';
            else if (pRaw === 'high') priority = 'High';
            else if (pRaw === 'critical') priority = 'Critical';

            let skills: string[] = [];
            const rawSkills = parsed.RequiredSkills || parsed.skills || parsed.requiredSkills;
            if (Array.isArray(rawSkills)) {
              skills = rawSkills.map((s: any) => {
                const match = SKILL_OPTIONS.find((opt) => opt.toLowerCase() === String(s).toLowerCase());
                return match || String(s);
              });
            }
            const projectDescription = String(
              parsed.projectDescription ||
              parsed.Projectdescription ||
              parsed.projectdescription ||
              ''
            );
            const architectNotes = String(
              parsed.architectNotes ||
              parsed.Projectdescription ||
              parsed.projectdescription ||
              projectDescription
            );

            setIntake({
              projectName,
              clientName,
              budget,
              startDate,
              endDate,
              projectDescription,
              priority,
              skills,
              architectNotes
            });
            setSourcePhase('complete');
          })
          .catch((err: any) => {
            console.error("AI upload failed:", err);
            setSourcePhase('complete');
          });
      },
        DOC_ANALYSIS_MS,
        DOC_ANALYSIS_AGENTS
      );
    } catch (err) {
      console.error("File analysis failed", err);
    }
  };

  const handlePlanChange = <K extends keyof PlanItem>(index: number, field: K, value: PlanItem[K]): void => {
    setPlanItems((prev) =>
      prev.map((row, i) => {
        if (i !== index) return row;

        const updated = { ...row, [field]: value } as PlanItem;
        const startIdx = Math.max(0, timelineColumns.indexOf(updated.startDate));
        const endIdx = Math.max(startIdx, timelineColumns.indexOf(updated.endDate));
        updated.startSlot = startIdx;
        updated.durationSlots = Math.max(1, endIdx - startIdx + 1);
        return updated;
      })
    );
  };

  const handleGroupSlNoChange = (oldVal: number, newVal: number): void => {
    setPlanItems((prev) =>
      prev.map((row) => (row.slNo === oldVal ? { ...row, slNo: newVal } : row))
    );
  };

  const handleGroupFeatureChange = (oldVal: string, newVal: string): void => {
    setPlanItems((prev) =>
      prev.map((row) => (row.feature === oldVal ? { ...row, feature: newVal } : row))
    );
  };

  const handleAddPlanRow = (): void => {
    setPlanItems((prev) => [
      ...prev,
      {
        slNo: prev.length > 0 ? prev[prev.length - 1].slNo + 1 : 1,
        feature: 'New Feature',
        task: 'New Task',
        description: '',
        startDate: timelineColumns[0] || 'New Date',
        endDate: timelineColumns[0] || 'New Date',
        efforts: 1,
        lane: 'development',
        startSlot: 0,
        durationSlots: 1,
        assignee: recommendedStaff[0]?.name ?? 'Unassigned',
        status: 'Not Started'
      }
    ]);
  };


  const handleAddTimelineColumn = (): void => {
    const nextLabel = `Day-${timelineColumns.length + 1}`;
    setTimelineColumns((prev) => [...prev, nextLabel]);
  };

  const handleRemoveTimelineColumn = (): void => {
    setTimelineColumns((prev) => {
      if (prev.length <= 1) return prev;
      const next = prev.slice(0, -1);
      setPlanItems((items) =>
        items.map((item) => {
          const maxSlot = next.length - 1;
          const startSlot = Math.min(item.startSlot, maxSlot);
          const durationSlots = Math.max(1, Math.min(item.durationSlots, next.length - startSlot));
          return {
            ...item,
            startSlot,
            durationSlots,
            startDate: next[startSlot],
            endDate: next[Math.min(next.length - 1, startSlot + durationSlots - 1)]
          };
        })
      );
      return next;
    });
  };

  const handleRenameTimelineColumn = (index: number, label: string): void => {
    const trimmed = label.trim();
    if (!trimmed) return;
    setTimelineColumns((prev) => {
      const oldLabel = prev[index];
      const next = [...prev];
      next[index] = trimmed;
      setPlanItems((items) =>
        items.map((item) => ({
          ...item,
          startDate: item.startDate === oldLabel ? trimmed : item.startDate,
          endDate: item.endDate === oldLabel ? trimmed : item.endDate
        }))
      );
      return next;
    });
  };

  const handleDeletePlanRow = (index: number): void => {
    setPlanItems((prev) => prev.filter((_, i) => i !== index));
  };

  const handleExportProjectPlan = (): void => {
    const headers = [
      'Sl.No',
      'Feature',
      'Tasks',
      'Description',
      'Type',
      'Start Date',
      'End Date',
      'Effort (hrs)',
      'Assignee',
      'Status',
      ...timelineColumns
    ];

    const rows = planItems.map((item) => {
      const timelineCells = timelineColumns.map((_, slot) =>
        slot >= item.startSlot && slot < item.startSlot + item.durationSlots ? item.lane : ''
      );
      return [
        item.slNo,
        item.feature,
        item.task,
        item.description || '',
        item.lane,
        item.startDate,
        item.endDate,
        item.efforts,
        item.assignee,
        item.status,
        ...timelineCells
      ];
    });

    const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'Project Plan');
    const safeProjectName = (intake.projectName || 'Project Plan').replace(/[\\/:*?"<>|]+/g, '-').trim();
    XLSX.writeFile(workbook, `${safeProjectName} - Project Plan.xlsx`);
  };

  const renderStepper = (): JSX.Element => (
    <div className={styles.stepper}>
      {STEPS.map((label, index) => {
        const stepNum = index + 1;
        const nextStep = stepNum + 1;
        const isDone = stepNum < step;
        const isActive = stepNum === step && !aiLoading;
        const showCheck = isDone && stepNum !== step;
        const connectorDone = nextStep < step;
        const connectorToActive = nextStep === step;

        return (
          <div key={label} className={styles.stepItem}>
            {index < STEPS.length - 1 && (
              <div
                className={`${styles.stepConnector} ${connectorDone ? styles.connectorDone : ''} ${connectorToActive ? styles.connectorToActive : ''}`}
              />
            )}
            <div
              className={`${styles.stepCircle} ${isActive ? styles.active : ''} ${isDone ? styles.done : ''}`}
            >
              {showCheck ? <CheckIcon /> : stepNum}
            </div>
            <span
              className={`${styles.stepLabel} ${isActive ? styles.active : ''} ${isDone ? styles.done : ''}`}
            >
              {label}
            </span>
          </div>
        );
      })}
    </div>
  );

  const loadingPhases = loadingAgents.map((a) => a.task);
  const activePhaseIndex = Math.min(
    loadingPhases.length - 1,
    Math.floor((aiProgress / 100) * loadingPhases.length)
  );

  const renderAiLoading = (title: string, agents?: AgentDef[]): JSX.Element => {
    const agentList = agents ?? loadingAgents;
    const phases = agentList.map((a) => a.task);
    const statusText = aiMessage || phases[activePhaseIndex];
    return (
      <div className={styles.aiLoading}>
        <AiSparkleLoader />
        <h3 className={styles.aiLoadingTitle}>{title}</h3>
        <p className={styles.aiLoadingSub}>
          {statusText}
          <span className={styles.aiLoadingDots} aria-hidden="true">
            <span>.</span>
            <span>.</span>
            <span>.</span>
          </span>
        </p>
        <div className={styles.aiProgressWrap}>
          <div className={styles.aiProgress}>
            <div className={styles.aiProgressBar} style={{ width: `${aiProgress}%` }} />
          </div>
          <span className={styles.aiProgressPct}>{Math.round(aiProgress)}%</span>
        </div>

        <ul className={styles.aiAgentSteps} aria-label="Agent processing steps">
          {agentList.map((agent, idx) => {
            const isDone = idx < activePhaseIndex;
            const isActive = idx === activePhaseIndex;
            return (
              <li
                key={agent.id}
                className={`${styles.aiAgentStep} ${isDone ? styles.aiAgentStepDone : ''} ${isActive ? styles.aiAgentStepActive : ''}`}
              >
                <span className={styles.aiAgentStepMark} aria-hidden="true">
                  {isDone ? '✓' : isActive ? <span className={styles.aiAgentStepSpinner} /> : '○'}
                </span>
                <span className={styles.aiAgentStepLabel}>{agent.label}</span>
                {isActive && <span className={styles.aiAgentStepStatus}>processing</span>}
              </li>
            );
          })}
        </ul>
      </div>
    );
  };

  const renderSourceResult = (): JSX.Element => (
    <div className={styles.sourceResult}>
      <div className={styles.sourceResultBadge}>
        <CheckIcon />
        <span>Analysis complete</span>
      </div>
      <h3 className={styles.sourceResultTitle}>Agent extraction summary</h3>
      <p className={styles.sourceResultFile}>
        Source: <strong>{uploadedFileName}</strong>
      </p>
      <div className={styles.sourceResultCard}>
        <dl className={styles.sourceResultGrid}>
          <div>
            <dt>Project</dt>
            <dd>{intake.projectName}</dd>
          </div>
          <div>
            <dt>Client</dt>
            <dd>{intake.clientName}</dd>
          </div>
          <div>
            <dt>Budget</dt>
            <dd>{intake.budget}</dd>
          </div>
          <div>
            <dt>Priority</dt>
            <dd>{intake.priority}</dd>
          </div>
          <div className={styles.sourceResultFull}>
            <dt>Skills detected</dt>
            <dd>{intake.skills.length > 0 ? intake.skills.join(', ') : '—'}</dd>
          </div>
          <div className={styles.sourceResultFull}>
            <dt>Scope notes</dt>
            <dd>{intake.architectNotes || intake.projectDescription || '—'}</dd>
          </div>
        </dl>
      </div>
      <p className={styles.sourceResultHint}>
        Fields pre-filled in Intake Spec — advancing automatically…
      </p>
    </div>
  );

  const renderSourceUpload = (): JSX.Element => (
    <div className={styles.sourceUpload}>
      <button
        type="button"
        className={styles.uploadZone}
        onClick={() => sourceFileRef.current?.click()}
        aria-label="Upload document"
      >
        <span className={styles.uploadIcon} aria-hidden="true">
          <svg viewBox="0 0 24 24" width="32" height="32" fill="none">
            <path d="M12 16V4M12 4l-4 4M12 4l4 4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M4 14v4a2 2 0 002 2h12a2 2 0 002-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
          </svg>
        </span>
        <strong>Drop file here or click to browse</strong>
        <span>PDF, DOCX, or TXT — max 25 MB</span>
      </button>
      <input
        ref={sourceFileRef}
        type="file"
        className={styles.hiddenFileInput}
        accept=".pdf,.doc,.docx,.txt"
        onChange={handleSourceFileSelected}
        aria-hidden="true"
        tabIndex={-1}
      />
      <button type="button" className={styles.sourceBackLink} onClick={resetSourceStep}>
        ← Choose a different source method
      </button>
    </div>
  );

  const renderSourceHero = (title: string, desc: string, icon: React.ReactNode): JSX.Element => (
    <div className={styles.sourceHero}>
      <div className={styles.sourceHeroIcon}>{icon}</div>
      <div>
        <h2 className={styles.sourceHeroTitle}>{title}</h2>
        <p className={styles.sourceHeroDesc}>{desc}</p>
      </div>
    </div>
  );

  const renderSourceChoose = (): JSX.Element => (
    <div className={styles.sourcePanel}>
      <div className={styles.sourceOptions}>
        <button
          type="button"
          className={`${styles.sourceOptionCard} ${styles.sourceOptionUpload}`}
          onClick={handleSelectDocumentUpload}
        >
          <span className={styles.sourceOptionAccent} aria-hidden="true" />
          <span className={styles.sourceOptionBadge}>AI Powered</span>
          <span className={`${styles.sourceOptionIconWrap} ${styles.sourceOptionIconUpload}`}>
            <DocumentUploadIcon />
          </span>
          <span className={styles.sourceOptionTitle}>Document Upload</span>
          <span className={styles.sourceOptionDesc}>
            Upload a SOW or brief. The AI agent will analyze it and pre-fill intake fields.
          </span>
          <ul className={styles.sourceOptionFeatures}>
            <li>PDF, DOCX, TXT</li>
            <li>Auto-extract fields</li>
          </ul>
          <span className={styles.sourceOptionCta}>
            Select upload <CardArrowIcon />
          </span>
        </button>
        <button
          type="button"
          className={`${styles.sourceOptionCard} ${styles.sourceOptionManual}`}
          onClick={handleSelectManualEntry}
        >
          <span className={styles.sourceOptionAccent} aria-hidden="true" />
          <span className={`${styles.sourceOptionBadge} ${styles.sourceOptionBadgeMuted}`}>Self-guided</span>
          <span className={`${styles.sourceOptionIconWrap} ${styles.sourceOptionIconManual}`}>
            <ManualEntryIcon />
          </span>
          <span className={styles.sourceOptionTitle}>Manual Entry</span>
          <span className={styles.sourceOptionDesc}>
            Enter project specifications yourself on the next step.
          </span>
          <ul className={styles.sourceOptionFeatures}>
            <li>Full control</li>
            <li>Skip analysis</li>
          </ul>
          <span className={styles.sourceOptionCta}>
            Continue manually <CardArrowIcon />
          </span>
        </button>
      </div>
    </div>
  );

  const renderSourceStep = (): JSX.Element => {
    if (sourcePhase === 'analyzing' && aiLoading) {
      return (
        <div className={styles.sourceStepInner}>
          {renderAiLoading('Analyzing uploaded document', DOC_ANALYSIS_AGENTS)}
        </div>
      );
    }
    if (sourcePhase === 'complete') {
      return <div className={styles.sourceStepInner}>{renderSourceResult()}</div>;
    }
    if (sourceMode === 'upload' && sourcePhase === 'upload') {
      return (
        <div className={styles.sourceStepInner}>
          {renderSourceHero(
            'Document Upload',
            'Upload your statement of work or project brief. Our agent will extract key intake fields automatically.',
            <DocumentUploadIcon />
          )}
          {renderSourceUpload()}
        </div>
      );
    }
    return (
      <div className={styles.sourceStepInner}>
        {renderSourceHero(
          'Choose Intake Source',
          'Select how you want to provide project information for this intake wizard.',
          <SourceDiamondIcon />
        )}
        {renderSourceChoose()}
      </div>
    );
  };

  const renderIntakeSpec = (): JSX.Element => (
  <>
    <h2 className={styles.sectionTitle}>
      <span className={styles.sectionIcon} aria-hidden="true"></span>
      Project Specification Entry
    </h2>
    <p className={styles.sectionDesc}>
      Provide direct properties to formulate optimized AI capacity allocations.
    </p>
    <div className={styles.formGrid}>
      <div className={styles.field}>
        <label className={styles.label} htmlFor="projectName">Project Name</label>
        <input
          id="projectName"
          className={styles.input}
          value={intake.projectName}
          onChange={(e) => updateIntake({ projectName: e.target.value })}
        />
      </div>
      <div className={styles.field}>
        <label className={styles.label} htmlFor="clientName">Client Name</label>
        <input
          id="clientName"
          className={styles.input}
          value={intake.clientName}
          onChange={(e) => updateIntake({ clientName: e.target.value })}
        />
      </div>
      <div className={styles.field}>
        <label className={styles.label} htmlFor="budget">Budget Limit Amount</label>
        <input
          id="budget"
          className={styles.input}
          value={intake.budget}
          onChange={(e) => updateIntake({ budget: e.target.value })}
        />
      </div>
      <div className={styles.field}>
        <label className={styles.label} htmlFor="startDate">Start Date</label>
        <input
          id="startDate"
          type="date"
          className={styles.input}
          value={intake.startDate}
          onChange={(e) => updateIntake({ startDate: e.target.value })}
        />
      </div>
      <div className={styles.field}>
        <label className={styles.label} htmlFor="endDate">End Date</label>
        <input
          id="endDate"
          type="date"
          className={styles.input}
          value={intake.endDate}
          onChange={(e) => updateIntake({ endDate: e.target.value })}
        />
      </div>
      <div className={`${styles.field} ${styles.fieldFull}`}>
        <label className={styles.label} htmlFor="projectDescription">Project Description</label>
        <textarea
          id="projectDescription"
          className={`${styles.textarea} ${css.textareaCenterTop}`}
          rows={4}
          value={intake.projectDescription}
          onChange={(e) => updateIntake({ projectDescription: e.target.value })}
        />
      </div>
    </div>
    <div className={styles.priorityRow}>
      <span className={styles.label}>Select Priority Level</span>
      <div className={styles.priorityBtns}>
        {PRIORITIES.map((p) => (
          <button
            key={p}
            type="button"
            className={`${styles.priorityBtn} ${intake.priority === p ? styles.selected : ''}`}
            onClick={() => updateIntake({ priority: p })}
          >
            {p}
          </button>
        ))}
      </div>
    </div>
    <div className={styles.skillsSection}>
      <span className={styles.skillsLabel}>Required Skills</span>
      <div className={styles.skillsChips}>
        {SKILL_OPTIONS.map((skill) => {
          const selected = intake.skills.indexOf(skill) >= 0;
          return (
            <button
              key={skill}
              type="button"
              className={`${styles.skillChip} ${selected ? styles.selected : ''}`}
              onClick={() => toggleSkill(skill)}
            >
              {selected ? '✓' : '+'} {skill}
            </button>
          );
        })}
      </div>
    </div>
  </>
  );

  const renderStaffRecommendation = (): JSX.Element => {
    if (aiLoading) {
      return renderAiLoading('Recommended Resource Allocation');
    }

    const avgMatch = recommendedStaff.length
      ? Math.round(recommendedStaff.reduce((sum, member) => sum + member.match, 0) / recommendedStaff.length)
      : null;
    const riskLabel =
      intake.priority === 'Critical' || intake.priority === 'High'
        ? 'High Risk'
        : intake.priority === 'Medium'
          ? 'Medium Risk'
          : 'Low Risk';

    return (
      <>
        <h2 className={styles.sectionTitle}>Recommended Resource Allocation</h2>
        <p className={styles.sectionDesc}>
          Continuous cognitive mapping analyzed load metrics, historical success and client contract alignment.
        </p>
        <div className={styles.staffLayout}>
          <div className={styles.staffGrid}>
            {recommendedStaff.length === 0 && (
              <p className={styles.sectionDesc}>No staff recommendations yet. Continue from intake to run the AI agent.</p>
            )}
            {recommendedStaff.map((member) => (
              <div key={member.initials} className={styles.staffCard}>
                <div className={styles.staffCardTop}>
                  <span className={styles.roleBadge}>{member.role}</span>
                  <span className={styles.matchScore}>Match: {member.match}%</span>
                </div>
                <div className={styles.staffProfile}>
                  <div className={styles.avatar}>{member.initials}</div>
                  <div>
                    <p className={styles.staffName}>{member.name}</p>
                    <p className={styles.staffDept}>{member.department}</p>
                  </div>
                </div>
                <div className={styles.staffStats}>
                  <span>Availability: <strong>{member.availability}%</strong></span>
                  <span>Success: <strong>{member.success}%</strong></span>
                </div>
              </div>
            ))}
          </div>
          <aside className={styles.cognitivePanel}>
            <h3 className={styles.cognitiveTitle}>COGNITIVE OVERVIEW</h3>
            <div className={styles.cognitiveRow}>
              <span className={styles.cognitiveLabel}>AI Confidence Score</span>
              <span className={styles.confidenceBox}>{avgMatch !== null ? `${avgMatch}%` : '—'}</span>
            </div>
            <div className={styles.cognitiveRow}>
              <span className={styles.cognitiveLabel}>Estimated Budget</span>
              <span className={styles.cognitiveValue}>{intake.budget}</span>
            </div>
            <div className={styles.cognitiveRow}>
              <span className={styles.cognitiveLabel}>Start Date</span>
              <span className={styles.cognitiveValue}>{intake.startDate}</span>
            </div>
            <div className={styles.cognitiveRow}>
              <span className={styles.cognitiveLabel}>End Date</span>
              <span className={styles.cognitiveValue}>{intake.endDate}</span>
            </div>
            <div className={styles.cognitiveRow}>
              <span className={styles.cognitiveLabel}>Delivery Risk Level</span>
              <span className={styles.riskBadge}>{riskLabel}</span>
            </div>
          </aside>
        </div>
      </>
    );


    
  };

  const renderProjectPlan = (): JSX.Element => {
    if (aiLoading) {
      return renderAiLoading(PLAN_LOADING_TITLE, PLAN_GENERATION_AGENTS);
    }

    // Group items by feature
    const groups: Record<string, { items: PlanItem[]; indices: number[]; totalEfforts: number; groupNum: number }> = {};
    let groupCounter = 0;
    const orderedFeatures: string[] = [];

    planItems.forEach((item, index) => {
      const feat = item.feature || 'General';
      if (orderedFeatures.indexOf(feat) === -1) {
        orderedFeatures.push(feat);
        groupCounter++;
        groups[feat] = {
          items: [],
          indices: [],
          totalEfforts: 0,
          groupNum: groupCounter
        };
      }
      groups[feat].items.push(item);
      groups[feat].indices.push(index);
      groups[feat].totalEfforts += item.efforts;
    });

    const monthSpans = getMonthSpans(timelineColumns);

    const renderChevronDown = () => (
      <svg className={css.chevronIcon} width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
        <polyline points="6 9 12 15 18 9"></polyline>
      </svg>
    );

    const renderChevronRight = () => (
      <svg className={css.chevronIcon} width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
        <polyline points="9 18 15 12 9 6"></polyline>
      </svg>
    );

    const renderFolderIcon = () => (
      <svg className={css.folderIcon} width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
        <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
      </svg>
    );

    const totalEffort = planItems.reduce((sum, item) => sum + item.efforts, 0);
    const uniqueAssignees = new Set(planItems.map((item) => item.assignee).filter(Boolean));
    const durationDays = getDurationDays(intake.startDate, intake.endDate);
    const dateRangeLabel =
      timelineColumns.length > 0
        ? `${timelineColumns[0]} – ${timelineColumns[timelineColumns.length - 1]}`
        : 'No dates';
    const timeMarkerSlot = 0;
    const zoomScale = planZoom / 100;
    const memberCount = uniqueAssignees.size || recommendedStaff.length;
    const avgSuccess = recommendedStaff.length
      ? Math.round(recommendedStaff.reduce((sum, member) => sum + member.success, 0) / recommendedStaff.length)
      : null;

    if (planItems.length === 0) {
      return (
        <div className={styles.planStudio}>
          <h2 className={styles.planPageTitle}>Project Plan</h2>
          <p className={styles.planPageSubtitle}>Generate a plan from the staff recommendation step to populate the timeline.</p>
        </div>
      );
    }

    return (
      <div className={styles.planStudio}>
        <div className={styles.planPageHead}>
          <div>
            <h2 className={styles.planPageTitle}>Project Plan</h2>
            <p className={styles.planPageSubtitle}>
              {intake.projectName} · Cognitive timeline generated from intake and staff allocation
            </p>
          </div>
        </div>

        <section className={styles.planStatsRow} aria-label="Plan summary">
          <article className={`${styles.planStatCard} ${styles.planStatEffort}`}>
            <div className={styles.planStatCardBar} aria-hidden="true" />
            <div className={styles.planStatCardHead}>
              <span className={styles.planStatIcon} aria-hidden="true"><PlanEffortIcon /></span>
            </div>
            <p className={styles.planStatLabel}>Total Effort</p>
            <p className={styles.planStatValue}>{totalEffort} hrs</p>
            <span className={`${styles.planStatTrend} ${styles.planStatTrendPurple}`}>Planned Hours</span>
            <div className={`${styles.planStatWatermark} ${styles.planStatWatermarkPurple}`} aria-hidden="true">
              <PlanEffortIcon />
            </div>
          </article>
          <article className={`${styles.planStatCard} ${styles.planStatDuration}`}>
            <div className={styles.planStatCardBar} aria-hidden="true" />
            <div className={styles.planStatCardHead}>
              <span className={styles.planStatIcon} aria-hidden="true"><PlanDurationIcon /></span>
            </div>
            <p className={styles.planStatLabel}>Estimated Duration</p>
            <p className={styles.planStatValue}>{durationDays} Days</p>
            <span className={`${styles.planStatTrend} ${styles.planStatTrendBlue}`}>Timeline Span</span>
            <div className={`${styles.planStatWatermark} ${styles.planStatWatermarkBlue}`} aria-hidden="true">
              <PlanDurationIcon />
            </div>
          </article>
          <article className={`${styles.planStatCard} ${styles.planStatTeam}`}>
            <div className={styles.planStatCardBar} aria-hidden="true" />
            <div className={styles.planStatCardHead}>
              <span className={styles.planStatIcon} aria-hidden="true"><PlanTeamIcon /></span>
            </div>
            <p className={styles.planStatLabel}>Resources Allocated</p>
            <p className={styles.planStatValue}>{memberCount}</p>
            <span className={`${styles.planStatTrend} ${styles.planStatTrendGreen}`}>
              {memberCount === 1 ? '1 Member' : `${memberCount} Members`}
            </span>
            <div className={`${styles.planStatWatermark} ${styles.planStatWatermarkGreen}`} aria-hidden="true">
              <PlanTeamIcon />
            </div>
          </article>
        </section>

        <div className={styles.planToolbar}>
          <div className={styles.planToolbarLeft}>
            <button type="button" className={styles.planToolbarPrimary} onClick={handleAddPlanRow}>
              + Add Task
            </button>
            <button type="button" className={styles.planToolbarGhost} onClick={handleExportProjectPlan}>
              Export Excel
            </button>
          </div>
          <div className={styles.planToolbarRight}>
            <div className={styles.planZoomControl}>
              <button type="button" onClick={() => setPlanZoom((z) => Math.max(70, z - 10))} aria-label="Zoom out">−</button>
              <span>{planZoom}%</span>
              <button type="button" onClick={() => setPlanZoom((z) => Math.min(130, z + 10))} aria-label="Zoom in">+</button>
            </div>
            <span className={styles.planDateRange}>{dateRangeLabel}</span>
          </div>
        </div>

        <div className={styles.planWorkspace}>
          <div
            className={css.ganttWrap}
            style={{ ['--plan-zoom' as string]: zoomScale }}
          >
          <table className={css.ganttTable}>
            <thead>
            
            
              <tr className={css.timelineMonthRow}>
                <th rowSpan={3} className={css.indexColHeader}>#</th>
                <th rowSpan={3} className={css.featureColHeader}>Feature</th>
                <th rowSpan={3} className={css.taskColHeader}>Task</th>
                <th rowSpan={3} className={css.descColHeader}>Description</th>
                <th rowSpan={3} className={css.typeColHeader}>Type</th>
                <th rowSpan={3} className={css.dateColHeader}>Start Date</th>
                <th rowSpan={3} className={css.dateColHeader}>End Date</th>
                <th rowSpan={3} className={css.effortsColHeader}>Effort (hrs)</th>
                <th rowSpan={3} className={css.assigneeColHeader}>Assignee</th>
                <th rowSpan={3} className={css.statusColHeader}>Status</th>
                <th rowSpan={3} className={css.actionsColHeader}>Actions</th>
                {monthSpans.map((m, index) => (
                  <th key={`month-${index}`} colSpan={m.span} className={css.timelineMonthCol}>
                    {m.label}
                  </th>
                ))}
              </tr>

        
        
              <tr className={css.timelineNumRow}>
                {timelineColumns.map((column, index) => {
                  const parts = column.split('-');
                  const dateNum = parts[0] || column;
                  const isActive = column === timelineColumns[0];
                  return (
                    <th key={`num-${column}-${index}`} className={`${css.timelineNumCol} ${isActive ? css.activeDate : ''}`}>
                      {dateNum}
                    </th>
                  );
                })}
              </tr>

            
            
              <tr className={css.timelineDayRow}>
                {timelineColumns.map((column, index) => {
                  const isActive = column === timelineColumns[0];
                  return (
                    <th key={`day-${column}-${index}`} className={`${css.timelineDayCol} ${isActive ? css.activeDate : ''}`}>
                      {getWeekday(column)}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {orderedFeatures.map((feat) => {
                const group = groups[feat];
                const isExpanded = expandedFeatures[feat] !== false;

                return (
                  <React.Fragment key={`group-${feat}`}>
                 
                 
                    <tr className={css.groupHeaderRow}>
                      <td className={css.excelRowHeaderCell} onClick={() => toggleFeatureExpand(feat)}>
                        <span className={css.accordionTrigger}>
                          {isExpanded ? renderChevronDown() : renderChevronRight()}
                        </span>
                      </td>
                      <td className={css.groupFeatureNameCell}>
                        <div className={css.groupFeatureWrapper}>
                          <span className={css.folderIconWrapper}>{renderFolderIcon()}</span>
                          <input
                            className={styles.groupFeatureInput}
                            value={feat}
                            onChange={(e) => handleGroupFeatureChange(feat, e.target.value)}
                          />
                        </div>
                      </td>
                      <td className={css.groupEffortsSumCell} colSpan={9}>
                        <span className={css.groupEffortsText}>{group.totalEfforts} hrs</span>
                      </td>
                    
                    
                      {timelineColumns.map((_, slot) => (
                        <td key={`header-slot-${slot}`} className={css.timelineCell} />
                      ))}
                    </tr>

                    {/* Child Task Rows */}
                    {isExpanded &&
                      group.items.map((item, childIndex) => {
                        const origIndex = group.indices[childIndex];
                        const childNum = `${group.groupNum}.${childIndex + 1}`;

                        return (
                          <tr key={`${item.feature}-${item.task}-${origIndex}`} className={css.excelRow}>
                          
                          
                            <td className={css.decimalNumCell}>{childNum}</td>

                           
                           
                            <td className={css.blankFeatureCell} />

                     
                     
                            <td className={css.excelTaskCell}>
                              <input
                                className={styles.planInput}
                                value={item.task}
                                onChange={(e) => handlePlanChange(origIndex, 'task', e.target.value)}
                              />
                            </td>

                           
                           
                            <td className={css.excelDescCell}>
                              <textarea
                                className={styles.planTextarea}
                                value={item.description || ''}
                                onChange={(e) => handlePlanChange(origIndex, 'description', e.target.value)}
                                rows={1}
                              />
                            </td>

                           
                           
                            <td className={css.typeBadgeCell}>
                              <select
                                className={`${css.badgeSelect} ${
                                  item.lane === 'development'
                                    ? css.badgeDevelopment
                                    : item.lane === 'testing'
                                      ? css.badgeTesting
                                      : css.badgeDeployment
                                }`}
                                value={item.lane}
                                onChange={(e) => handlePlanChange(origIndex, 'lane', e.target.value as PlanLane)}
                              >
                                <option value="development">Development</option>
                                <option value="testing">Testing</option>
                                <option value="demo">Deployment</option>
                              </select>
                            </td>

                            {/* Start Date */}
                            <td>
                              <select
                                className={styles.planSelect}
                                value={item.startDate}
                                onChange={(e) => handlePlanChange(origIndex, 'startDate', e.target.value)}
                              >
                                {timelineColumns.map((column) => (
                                  <option key={`start-${column}`} value={column}>
                                    {column}
                                  </option>
                                ))}
                              </select>
                            </td>

                           
                            <td>
                              <select
                                className={styles.planSelect}
                                value={item.endDate}
                                onChange={(e) => handlePlanChange(origIndex, 'endDate', e.target.value)}
                              >
                                {timelineColumns.map((column) => (
                                  <option key={`end-${column}`} value={column}>
                                    {column}
                                  </option>
                                ))}
                              </select>
                            </td>

                            {/* Efforts */}
                            <td>
                              <input
                                className={styles.planInputCompact}
                                value={item.efforts}
                                type="number"
                                min={1}
                                onChange={(e) => handlePlanChange(origIndex, 'efforts', Math.max(1, Number(e.target.value) || 1))}
                              />
                            </td>

                            <td className={css.assigneeCell}>
                              <div className={styles.assigneeCellInner}>
                                <span className={styles.assigneeAvatar}>{getAssigneeInitials(item.assignee, recommendedStaff)}</span>
                                <select
                                  className={styles.planSelectAssignee}
                                  value={item.assignee}
                                  onChange={(e) => handlePlanChange(origIndex, 'assignee', e.target.value)}
                                >
                                  {recommendedStaff.map((member) => (
                                    <option key={member.name} value={member.name}>
                                      {member.name}
                                    </option>
                                  ))}
                                </select>
                              </div>
                            </td>

                            <td className={css.statusCell}>
                              <select
                                className={`${styles.planStatusSelect} ${
                                  item.status === 'Done'
                                    ? styles.statusDone
                                    : item.status === 'In Progress'
                                      ? styles.statusInProgress
                                      : styles.statusNotStarted
                                }`}
                                value={item.status}
                                onChange={(e) => handlePlanChange(origIndex, 'status', e.target.value as PlanStatus)}
                              >
                                {PLAN_STATUSES.map((status) => (
                                  <option key={status} value={status}>
                                    {status}
                                  </option>
                                ))}
                              </select>
                            </td>

            
                            <td className={css.excelActionCell}>
                              <div className={css.actionsContainer}>
                                <span className={css.threeDots}>⋮</span>
                                <button
                                  type="button"
                                  className={styles.planDeleteBtn}
                                  onClick={() => handleDeletePlanRow(origIndex)}
                                >
                                  Delete
                                </button>
                              </div>
                            </td>

                            {/* Capsule Timeline Pills */}
                            {timelineColumns.map((_, slot) => {
                              const filled = slot >= item.startSlot && slot < item.startSlot + item.durationSlots;
                              let pillClass = '';
                              if (filled) {
                                if (item.durationSlots === 1) {
                                  pillClass = css.timelinePillSingle;
                                } else if (slot === item.startSlot) {
                                  pillClass = css.timelinePillStart;
                                } else if (slot === item.startSlot + item.durationSlots - 1) {
                                  pillClass = css.timelinePillEnd;
                                } else {
                                  pillClass = css.timelinePillMiddle;
                                }
                              }

                              const laneColorClass =
                                item.lane === 'development'
                                  ? css.barDevelopment
                                  : item.lane === 'testing'
                                    ? css.barTesting
                                    : css.barDeployment;

                              return (
                                <td
                                  key={`${item.task}-${slot}`}
                                  className={`${css.timelineCell} ${slot === timeMarkerSlot ? css.timeMarkerCol : ''}`}
                                >
                                  {slot === timeMarkerSlot && childIndex === 0 && (
                                    <span className={css.timeMarkerLabel}>TIME</span>
                                  )}
                                  {filled && (
                                    <div className={`${css.timelinePill} ${pillClass} ${laneColorClass}`} />
                                  )}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })}
                  </React.Fragment>
                );
              })}

              {/* Total Row */}
              <tr className={css.excelTotalRow}>
                <td className={css.excelTotalLabel} colSpan={7}>Total</td>
                <td className={css.excelTotalValue}>{totalEffort} hrs</td>
                <td colSpan={3} />
                <td colSpan={timelineColumns.length} />
              </tr>
            </tbody>
          </table>
          </div>

          <aside className={styles.planSidebar}>
            <section className={styles.planSidebarBlock}>
              <h3 className={styles.planSidebarTitle}>AI Insights</h3>
              <ul className={styles.planInsightList}>
                <li><span>Total Effort</span><strong>{totalEffort} hrs</strong></li>
                <li><span>Duration</span><strong>{durationDays} days</strong></li>
                <li><span>Resources</span><strong>{uniqueAssignees.size} members</strong></li>
                <li><span>Delivery Confidence</span><strong className={styles.planInsightGood}>{avgSuccess !== null ? `${avgSuccess}%` : '—'}</strong></li>
                <li><span>Risk Level</span><strong className={styles.planInsightGood}>{intake.priority}</strong></li>
              </ul>
            </section>
            <section className={styles.planSidebarBlock}>
              <h3 className={styles.planSidebarTitle}>Legend</h3>
              <div className={styles.planSidebarLegend}>
                <span><i className={`${css.legendSwatch} ${css.barDevelopment}`} />Development</span>
                <span><i className={`${css.legendSwatch} ${css.barTesting}`} />Testing</span>
                <span><i className={`${css.legendSwatch} ${css.barDeployment}`} />Deployment</span>
              </div>
            </section>
          </aside>
        </div>
      </div>
    );
  };

  const renderApprovals = (): JSX.Element => {
    const staffCount = recommendedStaff.length;
    const durationLabel = getDurationMonths(intake.startDate, intake.endDate);
    const displaySkills = intake.skills;
    const priorityPillClass =
      intake.priority === 'Critical' ? styles.priorityPillCritical
      : intake.priority === 'High' ? styles.priorityPillHigh
      : intake.priority === 'Medium' ? styles.priorityPillMedium
      : styles.priorityPillLow;

    return (
      <div className={styles.approvalStep}>
        <span className={styles.approvalKicker}>Recommendation</span>
        <h2 className={styles.approvalTitle}>Manager Review &amp; Project Authorization</h2>
        <p className={styles.approvalSubtitle}>
          Review project requirements, staffing recommendations, financial commitments, and delivery
          timelines before final approval.
        </p>

        <div className={styles.approvalOverview}>
          <article className={styles.approvalProjectCard}>
            {intake.priority && (
              <span className={styles.approvalProjectBadge}>{intake.priority} Priority</span>
            )}
            <h3 className={styles.approvalProjectName}>{intake.projectName}</h3>
            <p className={styles.approvalProjectDesc}>{intake.projectDescription}</p>
            <div className={styles.approvalTags}>
              {displaySkills.map((skill) => (
                <span key={skill} className={styles.approvalTag}>{skill}</span>
              ))}
            </div>
          </article>

          <div className={styles.approvalMetaGrid}>
            <div className={styles.approvalMetaCard}>
              <span className={styles.approvalMetaLabel}>Client</span>
              <strong className={styles.approvalMetaValue}>{intake.clientName}</strong>
            </div>
            <div className={styles.approvalMetaCard}>
              <span className={styles.approvalMetaLabel}>Staff Allocation</span>
              <strong className={styles.approvalMetaValue}>
                {staffCount} Member{staffCount !== 1 ? 's' : ''}
              </strong>
            </div>
            <div className={styles.approvalMetaCard}>
              <span className={styles.approvalMetaLabel}>Budget</span>
              <strong className={styles.approvalMetaValue}>{intake.budget}</strong>
            </div>
            <div className={styles.approvalMetaCard}>
              <span className={styles.approvalMetaLabel}>Duration</span>
              <strong className={styles.approvalMetaValue}>{durationLabel}</strong>
            </div>
            <div className={styles.approvalMetaCard}>
              <span className={styles.approvalMetaLabel}>Start Date</span>
              <strong className={styles.approvalMetaValue}>{formatDisplayDate(intake.startDate)}</strong>
            </div>
            <div className={styles.approvalMetaCard}>
              <span className={styles.approvalMetaLabel}>Priority</span>
              <strong className={styles.approvalMetaValue}>
                <span className={`${styles.priorityPill} ${priorityPillClass}`}>
                  {intake.priority} Priority
                </span>
              </strong>
            </div>
          </div>
        </div>

        <section className={styles.approvalBlock}>
          <label className={styles.approvalBlockLabel} htmlFor="architectNotes">
            Lead Architect Recommendation
          </label>
          <textarea
            id="architectNotes"
            className={styles.approvalTextarea}
            placeholder="Provide staffing rationale, project constraints, technical considerations, or delivery recommendations..."
            value={intake.architectNotes}
            onChange={(e) => updateIntake({ architectNotes: e.target.value })}
          />
        </section>
      </div>
    );
  };

  const renderSuccess = (): JSX.Element => (
    <div className={styles.successPanel}>
      <div className={styles.successHero} aria-hidden="true">
        <div className={styles.successHalo} />
        <div className={styles.successIcon}>
          <SuccessCheckIcon />
        </div>
      </div>
      <h2 className={styles.successTitle}>Submitted for Approval Successfully!</h2>
      <p className={styles.successDesc}>
        <strong>{intake.projectName}</strong> has been submitted for manager approval. Returning you
        to the dashboard to view the updated project list.
      </p>
    </div>
  );

  const renderBody = (): JSX.Element => {
    switch (step) {
      case 1:
        return renderSourceStep();
      case 2:
        return renderIntakeSpec();
      case 3:
        return renderStaffRecommendation();
      case 4:
        return renderProjectPlan();
      case 5:
        if (aiLoading) {
          return renderAiLoading(SUBMIT_LOADING_TITLE, SUBMIT_AGENTS);
        }
        return renderApprovals();
      case 6:
        return renderSuccess();
      default:
        return renderIntakeSpec();
    }
  };

  const renderFooter = (): JSX.Element | null => {
    if (step === 6) {
      return (
        <div className={styles.footerSuccess}>
          <button
            type="button"
            className={styles.btnDashboard}
            onClick={onSuccess ?? onClose}
          >
            <ExternalLinkIcon />
            Go to Project Dashboard
          </button>
        </div>
      );
    }

    if (step === 1) {
      return (
        <div className={styles.footer}>
          <button type="button" className={styles.btnBack} disabled>
            <span className={styles.btnText}>BACK</span>
          </button>
          <div className={styles.footerRight} />
        </div>
      );
    }

    if ((step === 3 || step === 5) && aiLoading) {
      return (
        <div className={styles.footer}>
          <button type="button" className={styles.btnBack} disabled>
            <span className={styles.btnText}>BACK</span>
          </button>
          <button type="button" className={styles.btnPrimary} disabled>
            <span className={styles.btnText}>
              {step === 5 ? 'SUBMITTING…' : 'GENERATE PROJECT PLAN'}
            </span>
          </button>
        </div>
      );
    }

    return (
      <div className={styles.footer}>
        <button
          type="button"
          className={styles.btnBack}
          onClick={handleBack}
          disabled={step <= 1 || aiLoading}
        >
          <span className={styles.btnText}>BACK</span>
        </button>
        <div className={styles.footerRight}>
          {step === 5 && (
            <button type="button" className={styles.btnDraft}>
              <span className={styles.btnText}>Save Draft</span>
            </button>
          )}
          {step === 2 && (
            <button type="button" className={styles.btnPrimary} onClick={handleContinueFromIntake}>
              <span className={styles.btnText}>CONTINUE</span>
            </button>
          )}
          {step === 3 && !aiLoading && (
            <button type="button" className={styles.btnPrimary} onClick={handleGeneratePlan}>
              <span className={styles.btnText}>GENERATE PROJECT PLAN</span>
            </button>
          )}
          {step === 4 && !aiLoading && (
            <button type="button" className={styles.btnPrimary} onClick={handleSendForApproval}>
              <span className={styles.btnText}>REVIEW</span>
            </button>
          )}
          {step === 5 && !aiLoading && (
            <button type="button" className={styles.btnPrimary} onClick={() => { void handleApprove(); }}>
              <span className={styles.btnText}>Submit for Approval</span>
            </button>
          )}
        </div>
      </div>
    );
  };

  return (
    <div className={`${styles.wizard} ${step === 4 ? styles.wizardPlan : ''}`}>
      <div className={styles.card}>
        <header className={styles.header}>
          <div className={styles.headerLeft}>
            <div className={styles.logo}>
              <ChipIcon />
            </div>
            <h1 className={styles.title}>AI PROJECT INTAKE WIZARD</h1>
          </div>
          <button type="button" className={styles.closeBtn} aria-label="Close wizard" onClick={onClose}>
            ×
          </button>
        </header>

        {renderStepper()}

        <div className={`${styles.body} ${step === 1 ? styles.bodySource : ''}`}>{renderBody()}</div>

        {renderFooter()}
      </div>
    </div>
  );
};

export default ProjectIntakeForm;
