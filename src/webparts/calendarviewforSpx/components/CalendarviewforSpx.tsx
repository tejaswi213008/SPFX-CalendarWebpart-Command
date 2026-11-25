import * as React from 'react';
import { useState, useEffect, useCallback, useRef } from 'react';
import type { ICalendarviewforSpxProps } from './ICalendarviewforSpxProps';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './CalendarviewforSpx.module.scss';
import { spfi, SPFx } from '@pnp/sp';
import ReactQuill from 'react-quill';
import 'react-quill/dist/quill.snow.css';
import {
  Stack,
  Text,
  Dropdown,
  type IDropdownOption,
  IconButton,
  PrimaryButton,
  DefaultButton,
  Pivot,
  PivotItem,
  Label,
  TextField
} from '@fluentui/react';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/fields';
import '@pnp/sp/attachments';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/items/list';
import '@pnp/sp/lists/web';

export interface IConsultingItem {
  Id: number;
  Title?: string;
  FirstName?: string;
  PrenatalConsultDate?: string | undefined;
  ScheduledDateDelivery?: string | undefined;
  EDC?: string | undefined;
  MaternalDOB?: string | undefined;
  MaternalMRN?: string | undefined;
  AdditionalConsulInfo?: string | undefined;
  PregnancyStatus?: string | undefined;
  FetalDiagnoses?: string | undefined;
  PrimaryOrganSystem?: string | undefined;
  SecondaryOrganSystem?: string | undefined;
  PrenatalImaging?: string | undefined;
  GeneticTesting?: string | undefined;
  DeliveryMode?: string | undefined;
  DeliveryLocation?: string | undefined;
  DeliveryRoomConsid?: string | undefined;
  PGEDependent?: string | undefined;
  PostnatalCarePlans?: string | undefined;
  FamilySocialConsiderations?: string | undefined;
  ActiveConcernsPendingIssues?: string | undefined;
  PrenatalTeamDiscussion?: string | undefined;
  MultiTeamDiscussion?: string | undefined;
  BabyDOB?: string | undefined;
  BabyMRN?: string | undefined;
  Baby2MRN?: string | undefined;
  Baby3MRN?: string | undefined;
  Baby4MRN?: string | undefined;
}

export interface ICalendarEvent {
  id: string;
  itemId: number;
  title: string;
  date: Date;
  type: 'Consulting' | 'Scheduled';
}

const WEEKDAY_LABELS: string[] = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const MONTH_NAMES: string[] = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

const SITE_TIME_ZONE = 'America/New_York';

const RICH_TEXT_TOOLBAR_MODULES = {
  toolbar: [
    [{ font: [] }],
    [{ size: [] }],
    [{ header: [1, 2, 3, 4, 5, 6, false] }],
    ['bold', 'italic', 'underline', 'strike'],
    [{ color: [] }, { background: [] }],
    [{ script: 'sub' }, { script: 'super' }],
    [{ align: [] }],
    [{ list: 'ordered' }, { list: 'bullet' }],
    [{ indent: '-1' }, { indent: '+1' }],
    ['blockquote', 'code-block'],
    ['link'],
    ['clean']
  ]
};

const RICH_TEXT_FORMATS = [
  'font',
  'size',
  'header',
  'bold',
  'italic',
  'underline',
  'strike',
  'color',
  'background',
  'script',
  'align',
  'list',
  'indent',
  'blockquote',
  'code-block',
  'link'
];

const RICH_TEXT_READONLY_MODULES = { toolbar: false };

function RichTextEditor(props: { value: string; onChange: (value: string) => void; readOnly?: boolean; tabIndex?: number }): React.ReactElement {
  const value = props.value ?? '';
  const readOnly = props.readOnly === true;
  return (
    <div className={(styles as Record<string, string>).quillRte}>
      <ReactQuill
        theme="snow"
        value={value}
        onChange={(content: string): void => {
          if (!readOnly) props.onChange(content);
        }}
        modules={readOnly ? RICH_TEXT_READONLY_MODULES : RICH_TEXT_TOOLBAR_MODULES}
        formats={RICH_TEXT_FORMATS}
        readOnly={readOnly}
        placeholder=""
        tabIndex={props.tabIndex}
      />
    </div>
  );
}

function parseSharePointDate(value: string | undefined): Date | undefined {
  if (value === undefined || String(value).trim() === '') return undefined;
  const s = String(value).trim();
  const match = /^\/Date\((-?\d+)([+-]\d{4})?\)\/$/.exec(s);
  if (match) {
    const ticks = parseInt(match[1], 10);
    if (isNaN(ticks)) return undefined;
    const utcDate = new Date(ticks);
    return isNaN(utcDate.getTime()) ? undefined : utcDate;
  }
  const utc = new Date(s);
  if (isNaN(utc.getTime())) return undefined;

  try {
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: SITE_TIME_ZONE,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    });

    const parts = formatter.formatToParts(utc);
    const pick = (type: Intl.DateTimeFormatPartTypes): number =>
      Number(parts.find(p => p.type === type)?.value ?? '');

    const year = pick('year');
    const month = pick('month');
    const day = pick('day');
    const hour = pick('hour');
    const minute = pick('minute');
    const second = pick('second');

    if (!year || !month || !day || isNaN(hour) || isNaN(minute) || isNaN(second)) {
      return utc;
    }

    const local = new Date(year, month - 1, day, hour, minute, second);
    return isNaN(local.getTime()) ? utc : local;
  } catch {
    return utc;
  }
}


function getDaysInMonth(year: number, month: number): Date[] {
  const first = new Date(year, month, 1);
  const last = new Date(year, month + 1, 0);
  const startPad = first.getDay();
  const days: Date[] = [];
  const prevMonth = new Date(year, month, 0);
  const prevCount = prevMonth.getDate();

  for (let i = 0; i < startPad; i++) {
    days.push(new Date(year, month - 1, prevCount - startPad + i + 1));
  }
  for (let d = 1; d <= last.getDate(); d++) {
    days.push(new Date(year, month, d));
  }
  const remaining = 42 - days.length;
  for (let i = 1; i <= remaining; i++) {
    days.push(new Date(year, month + 1, i));
  }
  return days;
}

export type EventFilterType = 'all' | 'Consulting' | 'Scheduled';

export type ItemModalMode = 'new' | 'edit' | 'view';

export type FormTabId = 'consult' | 'medical' | 'administrative';

const HOUR_OPTIONS: IDropdownOption[] = Array.from({ length: 24 }, (_, h) => ({
  key: h,
  text: h.toString().padStart(2, '0')
}));

const MINUTE_OPTIONS: IDropdownOption[] = Array.from({ length: 61 }, (_, m) => ({
  key: m,
  text: m.toString().padStart(2, '0')
}));

export interface IItemFormData {
  Title: string;
  FirstName: string;
  PrenatalConsultDate: string;
  PrenatalConsultHour?: number;
  PrenatalConsultMinute?: number;
  ScheduledDateDelivery: string;
  ScheduledHour?: number;
  ScheduledMinute?: number;
  EDC: string;
  MaternalDOB: string;
  MaternalMRN: string;
  AdditionalConsulInfo: string;
  PregnancyStatus: string;
  FetalDiagnoses: string;
  PrimaryOrganSystem: string;
  SecondaryOrganSystem: string;
  PrenatalImaging: string;
  GeneticTesting: string;
  DeliveryMode: string;
  DeliveryLocation: string;
  DeliveryRoomConsid: string;
  PGEDependent: string;
  PostnatalCarePlans: string;
  FamilySocialConsiderations: string;
  ActiveConcernsPendingIssues: string;
  PrenatalTeamDiscussion: string;
  MultiTeamDiscussion: string;
  BabyDOB: string;
  BabyMRN: string;
  Baby2MRN: string;
  Baby3MRN: string;
  Baby4MRN: string;
}



function formatDateForInput(value: string | undefined): string {
  if (!value) return '';
  const d = parseSharePointDate(value);
  if (!d) return '';
  const y = d.getFullYear();
  const m = (`0${d.getMonth() + 1}`).slice(-2);
  const day = (`0${d.getDate()}`).slice(-2);
  return `${y}-${m}-${day}`;
}

function toSharePointDate(dateStr: string, hour = 12, minute = 0): string | undefined {
  if (!dateStr || dateStr.trim() === '') return undefined;
  const parts = dateStr.trim().split('-');
  if (parts.length < 3) return undefined;
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  if (!year || !month || !day) return undefined;
  const h = Math.max(0, Math.min(23, hour));
  const m = Math.max(0, Math.min(59, minute));
  const yStr = year.toString().padStart(4, '0');
  const mStr = month.toString().padStart(2, '0');
  const dStr = day.toString().padStart(2, '0');
  const hStr = h.toString().padStart(2, '0');
  const minStr = m.toString().padStart(2, '0');
  return `${yStr}-${mStr}-${dStr}T${hStr}:${minStr}:00`;
}

function sanitizeSharePointFileName(name: string): string {
  const raw = String(name ?? '').trim();
  if (!raw) return 'file';
  // SharePoint disallows: " * : < > ? / \ | and some URL-ish chars can be problematic
  const cleaned = raw
    .replace(/["*:<>?/\\|]/g, '_')
    .replace(/[~#%&{}+]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
  return cleaned.length ? cleaned : 'file';
}

type ChoiceField = {
  InternalName?: string;
  Choices?: string[] | { results?: string[] };
};

export default function CalendarviewforSpx(props: ICalendarviewforSpxProps): React.ReactElement<ICalendarviewforSpxProps> {
  const { context, listTitle, siteUrl, listInternalName, hasTeamsContext } = props;
  const sp = React.useMemo(() => spfi().using(SPFx(context)), [context]);
  const [events, setEvents] = useState<ICalendarEvent[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const [viewDate, setViewDate] = useState<Date>(() => new Date());
  const [dateFilter, setDateFilter] = useState<EventFilterType>('all');
  const [modalMode, setModalMode] = useState<ItemModalMode | null>(null);
  const [modalItemId, setModalItemId] = useState<number | null>(null);
  const [activeFormTab, setActiveFormTab] = useState<FormTabId>('consult');
  const [formData, setFormData] = useState<IItemFormData>({
    Title: '',
    FirstName: '',
    PrenatalConsultDate: '',
    PrenatalConsultHour: 0,
    PrenatalConsultMinute: 0,
    ScheduledDateDelivery: '',
    ScheduledHour: 0,
    ScheduledMinute: 0,
    EDC: '',
    MaternalDOB: '',
    MaternalMRN: '',
    AdditionalConsulInfo: '',
    PregnancyStatus: '',
    FetalDiagnoses: '',
    PrimaryOrganSystem: '',
    SecondaryOrganSystem: '',
    PrenatalImaging: '',
    GeneticTesting: '',
    DeliveryMode: '',
    DeliveryLocation: '',
    DeliveryRoomConsid: '',
    PGEDependent: '',
    PostnatalCarePlans: '',
    FamilySocialConsiderations: '',
    ActiveConcernsPendingIssues: '',
    PrenatalTeamDiscussion: '',
    MultiTeamDiscussion: '',
    BabyDOB: '',
    BabyMRN: '',
    Baby2MRN: '',
    Baby3MRN: '',
    Baby4MRN: ''
  });
  const [formLoading, setFormLoading] = useState<boolean>(false);
  const [formError, setFormError] = useState<string | null>(null);
  const [lastNameError, setLastNameError] = useState<string | undefined>(undefined);
  const [choiceOptions, setChoiceOptions] = useState<Record<string, string[]>>({});
  const [consultExistingAttachments, setConsultExistingAttachments] = useState<{ name: string; url: string }[]>([]);
  const [consultPendingFiles, setConsultPendingFiles] = useState<File[]>([]);
  const consultPendingFilesRef = useRef<File[]>([]);
  const consultFileInputRef = useRef<HTMLInputElement | null>(null);

  const getSanitizedChoices = useCallback((fieldName: string): string[] => {
    const seen = new Set<string>();
    const values = choiceOptions[fieldName] || [];
    const cleaned: string[] = [];
    values.forEach((v: string) => {
      const value = String(v ?? '').trim();
      if (!value) return;
      if (/^please\s*select\.{0,3}$/i.test(value)) return;
      const key = value.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      cleaned.push(value);
    });
    return cleaned;
  }, [choiceOptions]);

  const buildItemPayload = useCallback((payload: IItemFormData): Record<string, string | null> => ({
    Title: payload.Title?.trim() || '',
    FirstName: payload.FirstName?.trim() || '',
    PrenatalConsultDate: toSharePointDate(
      payload.PrenatalConsultDate,
      payload.PrenatalConsultHour ?? 0,
      payload.PrenatalConsultMinute ?? 0
    ) || null,
    ScheduledDateDelivery: toSharePointDate(
      payload.ScheduledDateDelivery,
      payload.ScheduledHour ?? 0,
      payload.ScheduledMinute ?? 0
    ) || null,
    EDC: toSharePointDate(payload.EDC) || null,
    MaternalDOB: toSharePointDate(payload.MaternalDOB) || null,
    MaternalMRN: payload.MaternalMRN || null,
    PregnancyStatus: payload.PregnancyStatus || null,
    PrimaryOrganSystem: payload.PrimaryOrganSystem || null,
    SecondaryOrganSystem: payload.SecondaryOrganSystem || null,
    DeliveryMode: payload.DeliveryMode || null,
    DeliveryLocation: payload.DeliveryLocation || null,
    PGEDependent: payload.PGEDependent || null,
    AdditionalConsulInfo: payload.AdditionalConsulInfo || null,
    FetalDiagnoses: payload.FetalDiagnoses || null,
    PrenatalImaging: payload.PrenatalImaging || null,
    GeneticTesting: payload.GeneticTesting || null,
    DeliveryRoomConsid: payload.DeliveryRoomConsid || null,
    PostnatalCarePlans: payload.PostnatalCarePlans || null,
    FamilySocialConsiderations: payload.FamilySocialConsiderations || null,
    ActiveConcernsPendingIssues: payload.ActiveConcernsPendingIssues || null,
    PrenatalTeamDiscussion: toSharePointDate(payload.PrenatalTeamDiscussion) || null,
    MultiTeamDiscussion: toSharePointDate(payload.MultiTeamDiscussion) || null,
    BabyDOB: toSharePointDate(payload.BabyDOB) || null,
    BabyMRN: payload.BabyMRN || null,
    Baby2MRN: payload.Baby2MRN || null,
    Baby3MRN: payload.Baby3MRN || null,
    Baby4MRN: payload.Baby4MRN || null
  }), []);

  const loadItems = useCallback(async (): Promise<void> => {
    if (!listTitle) {
      setLoading(false);
      setError('Please set the list title in web part properties.');
      return;
    }

    try {
      setLoading(true);
      setError(null);

      const items: IConsultingItem[] = await sp.web.lists
        .getByTitle(listTitle)
        .items.select(
          'Id',
          'Title',
          'PrenatalConsultDate',
          'ScheduledDateDelivery'
        )
        .top(5000)();
      const mapped: ICalendarEvent[] = [];

      const getDisplayName = (item: IConsultingItem): string => {
        const last = (item.Title ?? '').trim();
        return last || '(No name)';
      };

      items.forEach((item: IConsultingItem) => {
        const displayName = getDisplayName(item);

        const consultDate = parseSharePointDate(item.PrenatalConsultDate);
        if (consultDate) {
          mapped.push({
            id: `${item.Id}-consulting`,
            itemId: item.Id,
            title: displayName,
            date: consultDate,
            type: 'Consulting'
          });
        }

        const scheduledDate = parseSharePointDate(item.ScheduledDateDelivery);
        if (scheduledDate) {
          mapped.push({
            id: `${item.Id}-scheduled`,
            itemId: item.Id,
            title: displayName,
            date: scheduledDate,
            type: 'Scheduled'
          });
        }
      });

      setEvents(mapped);
    } catch (e) {
      const message = e instanceof Error ? e.message : 'Failed to load calendar';
      setError(message);
      setEvents([]);
    } finally {
      setLoading(false);
    }
  }, [sp, listTitle]);

  const baseUrl = (siteUrl && String(siteUrl).trim()) ? String(siteUrl).trim().replace(/\/$/, '') : (context?.pageContext?.web?.absoluteUrl || '');
  const listNameForForms = (listInternalName && String(listInternalName).trim()) ? String(listInternalName).trim() : listTitle;
   const getListNamesToTry = useCallback((): string[] => {
    const names = [listTitle];
    if (listTitle === 'SPFX-RMNHCalendar') names.push('SPFXRMNHCalendar');
    else if (listTitle === 'SPFXRMNHCalendar') names.push('SPFX-RMNHCalendar');
    else {
      const noHyphen = listTitle.replace(/-/g, '');
      if (noHyphen !== listTitle) names.push(noHyphen);
    }
    return names;
  }, [listTitle]);
  const fetchChoiceOptionsForList = useCallback(async (): Promise<void> => {
    if (!listTitle) return;
  
    try {
      const fields = (await sp.web.lists
        .getByTitle(listTitle)
        .fields
        .filter("TypeAsString eq 'Choice'")
        .select("InternalName", "Choices")()) as unknown as ChoiceField[];
  
      const choiceMap: Record<string, string[]> = {};
      fields.forEach((field: ChoiceField) => {
  
        if (field?.InternalName && field.Choices) {
  
          if (Array.isArray(field.Choices)) {
            choiceMap[field.InternalName] = field.Choices;
          }
  
          if (!Array.isArray(field.Choices) && Array.isArray(field.Choices?.results)) {
            choiceMap[field.InternalName] = field.Choices.results;
          }
  
        }
  
      });
  
      setChoiceOptions(choiceMap);
  
      console.log("Choice fields:", choiceMap);
  
    } catch (error) {
      console.error("Error loading choice fields", error);
    }
  
  }, [sp, listTitle]);
  const fetchItemForModal = useCallback(async (itemId: number): Promise<IConsultingItem> => {
    if (!baseUrl || !context?.spHttpClient) throw new Error('Missing site or context');
    const select = 'Id,Title,FirstName,PrenatalConsultDate,ScheduledDateDelivery,EDC,MaternalDOB,MaternalMRN,AdditionalConsulInfo,PregnancyStatus,FetalDiagnoses,PrimaryOrganSystem,SecondaryOrganSystem,PrenatalImaging,GeneticTesting,DeliveryMode,DeliveryLocation,DeliveryRoomConsid,PGEDependent,PostnatalCarePlans,FamilySocialConsiderations,ActiveConcernsPendingIssues,PrenatalTeamDiscussion,MultiTeamDiscussion,BabyDOB,BabyMRN,Baby2MRN,Baby3MRN,Baby4MRN';
    const names = getListNamesToTry();
    for (const listName of names) {
      const url = `${baseUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})?$select=${encodeURIComponent(select)}`;
      const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (res.ok) {
        const item = await res.json();
        return item;
      }
    }
    throw new Error('Item not found');
  }, [baseUrl, context?.spHttpClient, getListNamesToTry]);

  const loadAllAttachments = useCallback(async (itemId: number): Promise<void> => {
    if (!listTitle) return;
    try {
      const files = await sp.web.lists
        .getByTitle(listTitle)
        .items.getById(itemId)
        .attachmentFiles();

      const all: { name: string; url: string }[] = (files || []).map(
        (f: { FileName: string; ServerRelativeUrl: string }) => ({
          name: f.FileName,
          url: f.ServerRelativeUrl
        })
      );
      setConsultExistingAttachments(all);
    } catch {
      setConsultExistingAttachments([]);
    }
  }, [sp, listTitle]);

  const createListItem = useCallback(async (payload: IItemFormData): Promise<number> => {
    if (!listTitle) throw new Error('Missing list title');

    const body = buildItemPayload(payload);
    const result = await sp.web.lists.getByTitle(listTitle).items.add(body);
  
    // 🔥 FIX (SAFE EXTRACTION)
    const itemId = result?.data?.Id ?? result?.Id;
  
    if (!itemId) {
      console.error("Full result:", result);
      throw new Error('Item created but ID not returned');
    }
  
    return itemId;
  }, [buildItemPayload, sp, listTitle]);

  const updateListItem = useCallback(async (itemId: number, payload: IItemFormData): Promise<void> => {
    if (!listTitle) throw new Error('Missing list title');
    const body = buildItemPayload(payload);

    await sp.web.lists
      .getByTitle(listTitle)
      .items.getById(itemId)
      .update(body);
  }, [buildItemPayload, sp, listTitle]);
  useEffect(() => {
    if (typeof document === 'undefined') {
      return;
    }
    const existing = document.head.querySelector('link[data-bootstrap-icons]');
    if (!existing) {
      const link = document.createElement('link');
      link.rel = 'stylesheet';
      link.href = 'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css';
      link.setAttribute('data-bootstrap-icons', 'true');
      document.head.appendChild(link);
    }
  }, []);

  useEffect(() => {
    loadItems().catch(() => {
      // loadItems already sets error state; avoid unhandled rejection lint
    });
  }, [loadItems]);

  useEffect(() => {
    if (listTitle && Object.keys(choiceOptions).length === 0) {
      fetchChoiceOptionsForList().catch(() => {
        // already logs; avoid unhandled rejection lint
      });
    }
  }, [modalMode, listTitle, choiceOptions, fetchChoiceOptionsForList]);

  const year = viewDate.getFullYear();
  const month = viewDate.getMonth();
  const days = getDaysInMonth(year, month);

  const getEventsForDay = (d: Date): ICalendarEvent[] => {
    const filtered = dateFilter === 'all' ? events : events.filter(e => e.type === dateFilter);
    return filtered.filter(e =>
      e.date.getDate() === d.getDate() &&
      e.date.getMonth() === d.getMonth() &&
      e.date.getFullYear() === d.getFullYear()
    );
  };

  const formatTimeForLabel = (d: Date): string => {
    const h = d.getHours().toString().padStart(2, '0');
    const m = d.getMinutes().toString().padStart(2, '0');
    return `${h}:${m}`;
  };

  const isCurrentMonth = (d: Date): boolean => d.getMonth() === month;
  const isToday = (d: Date): boolean => {
    const t = new Date();
    return d.getDate() === t.getDate() && d.getMonth() === t.getMonth() && d.getFullYear() === t.getFullYear();
  };

  const prevMonth = (): void => setViewDate(new Date(year, month - 1));
  const nextMonth = (): void => setViewDate(new Date(year, month + 1));

  const handleMonthChange = (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;
    const newMonth = Number(option.key);
    if (isNaN(newMonth)) return;
    setViewDate(prev => new Date(prev.getFullYear(), newMonth, 1));
  };

  const handleYearChange = (_: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;
    const newYear = Number(option.key);
    if (isNaN(newYear)) return;
    setViewDate(prev => new Date(newYear, prev.getMonth(), 1));
  };

  const monthOptions: IDropdownOption[] = MONTH_NAMES.map((name, idx) => ({
    key: idx,
    text: name
  }));

  const yearOptions: IDropdownOption[] = [];
  for (let y = year - 5; y <= year + 5; y++) {
    yearOptions.push({ key: y, text: `${y}` });
  }

  const consultTabOrder = {
    prenatalDate: 1,
    prenatalHour: 2,
    prenatalMinute: 3,
    edc: 4,
    lastName: 5,
    firstName: 6,
    maternalDob: 7,
    maternalMrn: 8,
    additionalConsultInfo: 9,
    pregnancyStatus: 10,
    chooseFile: 11
  } as const;

  const medicalTabOrder = {
    fetalDiagnoses: 1,
    primaryOrgan: 2,
    secondaryOrgan: 3,
    prenatalImaging: 4,
    geneticTesting: 5,
    scheduledDelivery: 6,
    deliveryMode: 7,
    deliveryLocation: 8,
    babyMrn: 9,
    baby2Mrn: 10,
    baby3Mrn: 11,
    baby4Mrn: 12,
    deliveryRoom: 13,
    pgeDependent: 14,
    postnatalCare: 15,
    familySocial: 16,
    activeConcerns: 17
  } as const;

  /* Administrative tab: explicit tab order (not left-to-right grid order) */
  const administrativeTabOrder = {
    prenatalTeam: 1,
    multiTeam: 2,
    babyDob: 3,
    babyMrn: 4,
    baby2Mrn: 5,
    baby3Mrn: 6,
    baby4Mrn: 7
  } as const;

  const listPath = baseUrl && listNameForForms ? `${baseUrl}/Lists/${encodeURIComponent(listNameForForms)}` : '';
  const newFormUrl = listPath ? `${listPath}/NewForm.aspx` : '';

  const openNewModal = (): void => {
    setFormError(null);
    setLastNameError(undefined);
    setConsultExistingAttachments([]);
    setConsultPendingFiles([]);
    consultPendingFilesRef.current = [];
    setFormData({
      Title: '',
      FirstName: '',
      PrenatalConsultDate: '',
      PrenatalConsultHour: 0,
      PrenatalConsultMinute: 0,
      ScheduledDateDelivery: '',
      ScheduledHour: 0,
      ScheduledMinute: 0,
      EDC: '',
      MaternalDOB: '',
      MaternalMRN: '',
      AdditionalConsulInfo: '',
      PregnancyStatus: '',
      FetalDiagnoses: '',
      PrimaryOrganSystem: '',
      SecondaryOrganSystem: '',
      PrenatalImaging: '',
      GeneticTesting: '',
      DeliveryMode: '',
      DeliveryLocation: '',
      DeliveryRoomConsid: '',
      PGEDependent: '',
      PostnatalCarePlans: '',
      FamilySocialConsiderations: '',
      ActiveConcernsPendingIssues: '',
      PrenatalTeamDiscussion: '',
      MultiTeamDiscussion: '',
      BabyDOB: '',
      BabyMRN: '',
      Baby2MRN: '',
      Baby3MRN: '',
      Baby4MRN: ''
    });
    setModalItemId(null);
    setModalMode('new');
    setActiveFormTab('consult');
    fetchChoiceOptionsForList().catch(() => {
      // ignore
    });
  };

  const openEditModal = (ev: ICalendarEvent): void => {
    setModalItemId(ev.itemId);
    setModalMode('edit');
    setFormError(null);
    setLastNameError(undefined);
    setFormLoading(true);
    setActiveFormTab('consult');
    setConsultExistingAttachments([]);
    setConsultPendingFiles([]);
    consultPendingFilesRef.current = [];
    fetchChoiceOptionsForList().catch(() => {
      // ignore
    });
    fetchItemForModal(ev.itemId)
      .then((item: IConsultingItem) => {
        const prenatalDateObj = item.PrenatalConsultDate ? parseSharePointDate(item.PrenatalConsultDate) : undefined;
        const scheduledDateObj = item.ScheduledDateDelivery ? parseSharePointDate(item.ScheduledDateDelivery) : undefined;
        setFormData({
          Title: (item.Title !== undefined && item.Title !== null) ? String(item.Title) : '',
          FirstName: (item.FirstName !== undefined && item.FirstName !== null) ? String(item.FirstName) : '',
          PrenatalConsultDate: formatDateForInput(item.PrenatalConsultDate),
          PrenatalConsultHour: prenatalDateObj ? prenatalDateObj.getHours() : 0,
          PrenatalConsultMinute: prenatalDateObj ? prenatalDateObj.getMinutes() : 0,
          ScheduledDateDelivery: formatDateForInput(item.ScheduledDateDelivery),
          ScheduledHour: scheduledDateObj ? scheduledDateObj.getHours() : 0,
          ScheduledMinute: scheduledDateObj ? scheduledDateObj.getMinutes() : 0,
          EDC: formatDateForInput(item.EDC),
          MaternalDOB: formatDateForInput(item.MaternalDOB),
          MaternalMRN: (item.MaternalMRN !== undefined && item.MaternalMRN !== null) ? String(item.MaternalMRN) : '',
          AdditionalConsulInfo: (item.AdditionalConsulInfo !== undefined && item.AdditionalConsulInfo !== null) ? String(item.AdditionalConsulInfo) : '',
          PregnancyStatus: (item.PregnancyStatus !== undefined && item.PregnancyStatus !== null) ? String(item.PregnancyStatus) : '',
          FetalDiagnoses: (item.FetalDiagnoses !== undefined && item.FetalDiagnoses !== null) ? String(item.FetalDiagnoses) : '',
          PrimaryOrganSystem: (item.PrimaryOrganSystem !== undefined && item.PrimaryOrganSystem !== null) ? String(item.PrimaryOrganSystem) : '',
          SecondaryOrganSystem: (item.SecondaryOrganSystem !== undefined && item.SecondaryOrganSystem !== null) ? String(item.SecondaryOrganSystem) : '',
          PrenatalImaging: (item.PrenatalImaging !== undefined && item.PrenatalImaging !== null) ? String(item.PrenatalImaging) : '',
          GeneticTesting: (item.GeneticTesting !== undefined && item.GeneticTesting !== null) ? String(item.GeneticTesting) : '',
          DeliveryMode: (item.DeliveryMode !== undefined && item.DeliveryMode !== null) ? String(item.DeliveryMode) : '',
          DeliveryLocation: (item.DeliveryLocation !== undefined && item.DeliveryLocation !== null) ? String(item.DeliveryLocation) : '',
          DeliveryRoomConsid: (item.DeliveryRoomConsid !== undefined && item.DeliveryRoomConsid !== null) ? String(item.DeliveryRoomConsid) : '',
          PGEDependent: (item.PGEDependent !== undefined && item.PGEDependent !== null) ? String(item.PGEDependent) : '',
          PostnatalCarePlans: (item.PostnatalCarePlans !== undefined && item.PostnatalCarePlans !== null) ? String(item.PostnatalCarePlans) : '',
          FamilySocialConsiderations: (item.FamilySocialConsiderations !== undefined && item.FamilySocialConsiderations !== null) ? String(item.FamilySocialConsiderations) : '',
          ActiveConcernsPendingIssues: (item.ActiveConcernsPendingIssues !== undefined && item.ActiveConcernsPendingIssues !== null) ? String(item.ActiveConcernsPendingIssues) : '',
          PrenatalTeamDiscussion: formatDateForInput(item.PrenatalTeamDiscussion),
          MultiTeamDiscussion: formatDateForInput(item.MultiTeamDiscussion),
          BabyDOB: formatDateForInput(item.BabyDOB),
          BabyMRN: (item.BabyMRN !== undefined && item.BabyMRN !== null) ? String(item.BabyMRN) : '',
          Baby2MRN: (item.Baby2MRN !== undefined && item.Baby2MRN !== null) ? String(item.Baby2MRN) : '',
          Baby3MRN: (item.Baby3MRN !== undefined && item.Baby3MRN !== null) ? String(item.Baby3MRN) : '',
          Baby4MRN: (item.Baby4MRN !== undefined && item.Baby4MRN !== null) ? String(item.Baby4MRN) : ''
        });
        setFormLoading(false);
        loadAllAttachments(ev.itemId).catch(() => {
          // ignore; attachments can be refreshed on next open
        });
      })
      .catch((e: unknown) => {
        setFormError(e instanceof Error ? e.message : 'Failed to load item');
        setFormLoading(false);
      });
  };

  const openViewModal = (ev: ICalendarEvent): void => {
    setModalItemId(ev.itemId);
    setModalMode('view');
    setFormError(null);
    setLastNameError(undefined);
    setFormLoading(true);
    setActiveFormTab('consult');
    setConsultExistingAttachments([]);
    setConsultPendingFiles([]);
    consultPendingFilesRef.current = [];
    fetchChoiceOptionsForList().catch(() => {
      // ignore
    });
    fetchItemForModal(ev.itemId)
      .then((item: IConsultingItem) => {
        const prenatalDateObj = item.PrenatalConsultDate ? parseSharePointDate(item.PrenatalConsultDate) : undefined;
        const scheduledDateObj = item.ScheduledDateDelivery ? parseSharePointDate(item.ScheduledDateDelivery) : undefined;
        setFormData({
          Title: (item.Title !== undefined && item.Title !== null) ? String(item.Title) : '',
          FirstName: (item.FirstName !== undefined && item.FirstName !== null) ? String(item.FirstName) : '',
          PrenatalConsultDate: formatDateForInput(item.PrenatalConsultDate),
          PrenatalConsultHour: prenatalDateObj ? prenatalDateObj.getHours() : 0,
          PrenatalConsultMinute: prenatalDateObj ? prenatalDateObj.getMinutes() : 0,
          ScheduledDateDelivery: formatDateForInput(item.ScheduledDateDelivery),
          ScheduledHour: scheduledDateObj ? scheduledDateObj.getHours() : 0,
          ScheduledMinute: scheduledDateObj ? scheduledDateObj.getMinutes() : 0,
          EDC: formatDateForInput(item.EDC),
          MaternalDOB: formatDateForInput(item.MaternalDOB),
          MaternalMRN: (item.MaternalMRN !== undefined && item.MaternalMRN !== null) ? String(item.MaternalMRN) : '',
          AdditionalConsulInfo: (item.AdditionalConsulInfo !== undefined && item.AdditionalConsulInfo !== null) ? String(item.AdditionalConsulInfo) : '',
          PregnancyStatus: (item.PregnancyStatus !== undefined && item.PregnancyStatus !== null) ? String(item.PregnancyStatus) : '',
          FetalDiagnoses: (item.FetalDiagnoses !== undefined && item.FetalDiagnoses !== null) ? String(item.FetalDiagnoses) : '',
          PrimaryOrganSystem: (item.PrimaryOrganSystem !== undefined && item.PrimaryOrganSystem !== null) ? String(item.PrimaryOrganSystem) : '',
          SecondaryOrganSystem: (item.SecondaryOrganSystem !== undefined && item.SecondaryOrganSystem !== null) ? String(item.SecondaryOrganSystem) : '',
          PrenatalImaging: (item.PrenatalImaging !== undefined && item.PrenatalImaging !== null) ? String(item.PrenatalImaging) : '',
          GeneticTesting: (item.GeneticTesting !== undefined && item.GeneticTesting !== null) ? String(item.GeneticTesting) : '',
          DeliveryMode: (item.DeliveryMode !== undefined && item.DeliveryMode !== null) ? String(item.DeliveryMode) : '',
          DeliveryLocation: (item.DeliveryLocation !== undefined && item.DeliveryLocation !== null) ? String(item.DeliveryLocation) : '',
          DeliveryRoomConsid: (item.DeliveryRoomConsid !== undefined && item.DeliveryRoomConsid !== null) ? String(item.DeliveryRoomConsid) : '',
          PGEDependent: (item.PGEDependent !== undefined && item.PGEDependent !== null) ? String(item.PGEDependent) : '',
          PostnatalCarePlans: (item.PostnatalCarePlans !== undefined && item.PostnatalCarePlans !== null) ? String(item.PostnatalCarePlans) : '',
          FamilySocialConsiderations: (item.FamilySocialConsiderations !== undefined && item.FamilySocialConsiderations !== null) ? String(item.FamilySocialConsiderations) : '',
          ActiveConcernsPendingIssues: (item.ActiveConcernsPendingIssues !== undefined && item.ActiveConcernsPendingIssues !== null) ? String(item.ActiveConcernsPendingIssues) : '',
          PrenatalTeamDiscussion: formatDateForInput(item.PrenatalTeamDiscussion),
          MultiTeamDiscussion: formatDateForInput(item.MultiTeamDiscussion),
          BabyDOB: formatDateForInput(item.BabyDOB),
          BabyMRN: (item.BabyMRN !== undefined && item.BabyMRN !== null) ? String(item.BabyMRN) : '',
          Baby2MRN: (item.Baby2MRN !== undefined && item.Baby2MRN !== null) ? String(item.Baby2MRN) : '',
          Baby3MRN: (item.Baby3MRN !== undefined && item.Baby3MRN !== null) ? String(item.Baby3MRN) : '',
          Baby4MRN: (item.Baby4MRN !== undefined && item.Baby4MRN !== null) ? String(item.Baby4MRN) : ''
        });
        setFormLoading(false);
        loadAllAttachments(ev.itemId).catch(() => {
          // ignore; attachments can be refreshed on next open
        });
      })
      .catch((e: unknown) => {
        setFormError(e instanceof Error ? e.message : 'Failed to load item');
        setFormLoading(false);
      });
  };

  const closeModal = (): void => {
    setModalMode(null);
    setModalItemId(null);
    setFormError(null);
    setLastNameError(undefined);
    setConsultPendingFiles([]);
    consultPendingFilesRef.current = [];
  };
  const uploadConsultAttachments = useCallback(async (itemId: number, files: File[]): Promise<void> => {
    if (!listTitle || files.length === 0) return;
  
    const item = sp.web.lists.getByTitle(listTitle).items.getById(itemId);
  
    for (const file of files) {
      const fileName = sanitizeSharePointFileName(file.name);
  
      try {
        // 🔥 Delete if already exists
        await item.attachmentFiles.getByName(fileName).delete();
      } catch {
        // ignore if not exists
      }
  
      await item.attachmentFiles.add(fileName, file);
    }
  
    setConsultPendingFiles([]);
    consultPendingFilesRef.current = [];
  
    await loadAllAttachments(itemId);
  }, [listTitle, loadAllAttachments, sp]);

  const deleteAttachment = useCallback(async (itemId: number, fileUrl: string): Promise<void> => {
    if (!listTitle) return;
    const fileName = fileUrl.split('/').pop();
    if (!fileName) return;

    await sp.web.lists
      .getByTitle(listTitle)
      .items.getById(itemId)
      .attachmentFiles.getByName(fileName)
      .delete();

    await loadAllAttachments(itemId);
  }, [listTitle, loadAllAttachments, sp]);

  const handleSaveNew = async (): Promise<void> => {
    setFormError(null);
    setLastNameError(undefined);
    if (!formData.Title?.trim()) {
      setLastNameError('Last Name is required.');
      return;
    }
    try {
      setFormLoading(true);

      const filesToUpload = [...consultPendingFiles];

      const newId = await createListItem(formData);
  
      // 🔥 Upload only if files exist
      if (filesToUpload.length > 0) {
        await uploadConsultAttachments(newId, filesToUpload);
      }
  
      closeModal();
      await loadItems();
  
    } catch (e) {
      setLastNameError(undefined);
      setFormError(e instanceof Error ? e.message : 'Save failed');
    } finally {
      setFormLoading(false);
    }
  };

  const handleSaveEdit = async (): Promise<void> => {
    if (modalItemId === null) return;

    setFormError(null);
    setLastNameError(undefined);
    if (!formData.Title?.trim()) {
      setLastNameError('Last Name is required.');
      return;
    }
    try {
      setFormLoading(true);

      const filesToUpload = [...consultPendingFiles];

      await updateListItem(modalItemId, formData);
  
      if (filesToUpload.length > 0) {
        await uploadConsultAttachments(modalItemId, filesToUpload);
      }
  
      closeModal();
      await loadItems();
  
    } catch (e) {
      setLastNameError(undefined);
      setFormError(e instanceof Error ? e.message : 'Update failed');
    } finally {
      setFormLoading(false);
    }
  };

  const handleAddNew = (): void => {
    openNewModal();
  };

  const handleEditEvent = (ev: ICalendarEvent): void => {
    openViewModal(ev);
    };

  if (!listTitle) {
    return (
      <section className={`${styles.calendarviewforSpx} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.message}>Configure the list title in web part properties. Use display name (e.g. SPFX-RMNHCalendar) or internal name (e.g. SPFXRMNHCalendar).</div>
      </section>
    );
  }

  if (loading) {
    return (
      <section className={`${styles.calendarviewforSpx} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.message}>Loading calendar…</div>
      </section>
    );
  }

  if (error) {
    return (
      <section className={`${styles.calendarviewforSpx} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.error}>
          <strong>Calendar could not load</strong>
          <p>{error}</p>
          <p style={{ fontSize: 12, marginTop: 8, opacity: 0.95 }}>
            List not found (404). We tried both <strong>display name</strong> (SPFX-RMNHCalendar) and <strong>internal name</strong> (SPFXRMNHCalendar). Confirm the list exists in Site Contents of this site.
          </p>
        </div>
      </section>
    );
  }

  return (
    <section className={`${styles.calendarviewforSpx} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.calendar}>
        <div className={styles.calendarHeader}>
          <div className={(styles as Record<string, string>).calendarHeaderInner}>
            <div className={(styles as Record<string, string>).calendarFilters}>
              <Dropdown
                ariaLabel="Select month"
                selectedKey={month}
                options={monthOptions}
                onChange={handleMonthChange}
                styles={{ dropdown: { width: 140 } }}
              />
              <Dropdown
                ariaLabel="Select year"
                selectedKey={year}
                options={yearOptions}
                onChange={handleYearChange}
                styles={{ dropdown: { width: 100 } }}
              />
            </div>
            <div className={(styles as Record<string, string>).calendarTitle}>
              <IconButton
                iconProps={{ iconName: 'ChevronLeft' }}
                ariaLabel="Previous month"
                onClick={prevMonth}
              />
              <Text variant="xLarge" className={styles.monthTitle}>
                {MONTH_NAMES[month]} {year}
              </Text>
              <IconButton
                iconProps={{ iconName: 'ChevronRight' }}
                ariaLabel="Next month"
                onClick={nextMonth}
              />
            </div>
            {newFormUrl && (
              <div className={(styles as Record<string, string>).calendarHeaderRight}>
                <PrimaryButton
                  text="New Event"
                  iconProps={{ iconName: 'Add' }}
                  onClick={handleAddNew}
                  ariaLabel="Add new event"
                  className={(styles as Record<string, string>).newEventButton}
                />
              </div>
            )}
          </div>
        </div>
        <div className={styles.weekdayRow}>
          {WEEKDAY_LABELS.map(label => (
            <div key={label} className={styles.weekdayCell}>{label}</div>
          ))}
        </div>
        <div className={styles.daysGrid}>
          {days.map((d, idx) => (
            <div
              key={`${d.getFullYear()}-${d.getMonth()}-${d.getDate()}-${idx}`}
              className={`${styles.dayCell} ${!isCurrentMonth(d) ? styles.otherMonth : ''} ${isToday(d) ? styles.today : ''}`}
              onDoubleClick={(): void => openNewModal()}
              role="button"
              tabIndex={0}
              onKeyDown={(e): void => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openNewModal(); } }}
              aria-label={`${d.getDate()} ${MONTH_NAMES[d.getMonth()]} ${d.getFullYear()}. Double-click to add event.`}
            >
              <span className={styles.dayNumber}>{d.getDate()}</span>
              <div className={styles.events}>
                {getEventsForDay(d).map(ev => {
                  const timePrefix = ev.type === 'Consulting' ? `${formatTimeForLabel(ev.date)} ` : '';
                  return (
                  <div
                    key={ev.id}
                    role="button"
                    tabIndex={0}
                    className={`${styles.eventPill} ${ev.type === 'Consulting' ? styles.consultingEvent : styles.scheduledEvent}`}
                    title={`${timePrefix}${ev.title} (${ev.type} Date) – Double-click to open`}
                    onDoubleClick={(e): void => { e.stopPropagation(); handleEditEvent(ev); }}
                    onKeyDown={(e): void => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handleEditEvent(ev); } }}
                  >
                    <i
                      className={`${ev.type === 'Consulting' ? 'bi bi-calendar-event' : 'bi bi-bag-heart-fill'} ${(styles as Record<string, string>).eventIcon}`}
                      aria-hidden="true"
                    />
                    {timePrefix}{ev.title}
                  </div>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
        <div className={styles.legend}>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" wrap>
            <DefaultButton
              text="All"
              iconProps={{ iconName: 'BulletedList' }}
              onClick={(): void => setDateFilter('all')}
              styles={{
                root: {
                  borderColor: '#6264a7',
                  background: dateFilter === 'all' ? '#6264a7' : 'transparent',
                  color: dateFilter === 'all' ? '#ffffff' : '#6264a7'
                },
                label: {
                  color: dateFilter === 'all' ? '#ffffff' : '#6264a7'
                },
                icon: {
                  color: dateFilter === 'all' ? '#ffffff' : '#6264a7',
                  fontSize: 10
                }
              }}
            />
            <DefaultButton
              text="Consult Date"
              iconProps={{ iconName: 'Calendar' }}
              onClick={(): void => setDateFilter(dateFilter === 'Consulting' ? 'all' : 'Consulting')}
              styles={{
                root: {
                  borderColor: '#bfbfbf',
                  background: dateFilter === 'Consulting' ? '#bfbfbf' : 'transparent',
                  color: '#000000'
                },
                label: {
                  color: '#000000'
                },
                icon: {
                  color: '#000000',
                  fontSize: 10
                }
              }}
            />
            <DefaultButton
              text="Schedule Delivery Date"
              iconProps={{ iconName: 'Medical' }}
              onClick={(): void => setDateFilter(dateFilter === 'Scheduled' ? 'all' : 'Scheduled')}
              styles={{
                root: {
                  borderColor: '#ffd6db',
                  background: dateFilter === 'Scheduled' ? '#ffd6db' : 'transparent',
                  color: '#000000'
                },
                label: {
                  color: '#000000'
                },
                icon: {
                  color: '#000000',
                  fontSize: 10
                }
              }}
            />
          </Stack>
        </div>
      </div>
      {modalMode !== null && (
        <div
          className={styles.powerAppsModalBackdrop}
          role="dialog"
          aria-modal="true"
          aria-labelledby="item-modal-title"
        >
          <div
            className={`${styles.powerAppsModal} ${modalMode === 'view' ? (styles as Record<string, string>).powerAppsModalView : ''}`}
            onClick={(e): void => { e.stopPropagation(); }}
          >
            <div className={styles.powerAppsModalHeader}>
              <div className={(styles as Record<string, string>).modalHeaderActions}>
                {(modalMode === 'new' || modalMode === 'edit') && (
                  <PrimaryButton
                    text="Save"
                    iconProps={{ iconName: 'Save' }}
                    onClick={(): void => {
                      if (modalMode === 'new') {
                        handleSaveNew().catch(() => {
                          // handled inside handleSaveNew
                        });
                      } else {
                        handleSaveEdit().catch(() => {
                          // handled inside handleSaveEdit
                        });
                      }
                    }}
                  />
                )}
                {modalMode === 'view' && modalItemId !== null && (
                  <DefaultButton
                    text="Edit all"
                    className={(styles as Record<string, string>).accentHeaderButton}
                    iconProps={{ iconName: 'Edit' }}
                    onClick={(): void => {
                      const ev = events.filter((evt: ICalendarEvent) => evt.itemId === modalItemId)[0];
                      if (ev) openEditModal(ev);
                    }}
                  />
                )}
                <DefaultButton
                  text="Cancel"
                  iconProps={{ iconName: 'Cancel' }}
                  onClick={closeModal}
                />
              </div>
            </div>
            <div className={styles.powerAppsModalBody}>
              <Pivot
                selectedKey={activeFormTab}
                onLinkClick={(item): void => {
                  if (!item) return;
                  setActiveFormTab(item.props.itemKey as FormTabId);
                }}
              >
                <PivotItem headerText="Consult" itemKey="consult" />
                <PivotItem headerText="Medical" itemKey="medical" />
                <PivotItem headerText="Administrative" itemKey="administrative" />
              </Pivot>
              <div className={styles.modalForm}>
                {formLoading && <p className={styles.message}>Loading…</p>}
                {!formLoading && activeFormTab === 'consult' && (() => {
                  const isReadOnly = modalMode === 'view';
                  return (
                  <div className={styles.consultFormEdit}>
                    <Stack tokens={{ childrenGap: 16 }}>
                      <Stack horizontal tokens={{ childrenGap: 16 }} className={styles.consultFormRow}>
                        <Stack.Item grow>
                          <div className={styles.modalFormField}>
                            <Label htmlFor="modal-prenatal">Prenatal Consult Date &amp; time</Label>
                            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                              <TextField
                                id="modal-prenatal"
                                type="date"
                                ariaLabel="Prenatal consult date"
                                tabIndex={consultTabOrder.prenatalDate}
                                value={formData.PrenatalConsultDate}
                                onChange={(_, newValue): void =>
                                  setFormData(prev => ({ ...prev, PrenatalConsultDate: newValue ?? '' }))
                                }
                                styles={{ root: { width: '100%' } }}
                                disabled={isReadOnly}
                                readOnly={isReadOnly}
                              />
                              <Dropdown
                                ariaLabel="Hour"
                                tabIndex={consultTabOrder.prenatalHour}
                                selectedKey={formData.PrenatalConsultHour ?? 0}
                                options={HOUR_OPTIONS}
                                onChange={(_, option): void => {
                                  if (!option) return;
                                  setFormData(prev => ({ ...prev, PrenatalConsultHour: Number(option.key) }));
                                }}
                                styles={{ dropdown: { width: 70 } }}
                                disabled={isReadOnly}
                              />
                              <Dropdown
                                ariaLabel="Minute"
                                tabIndex={consultTabOrder.prenatalMinute}
                                selectedKey={formData.PrenatalConsultMinute ?? 0}
                                options={MINUTE_OPTIONS}
                                onChange={(_, option): void => {
                                  if (!option) return;
                                  setFormData(prev => ({ ...prev, PrenatalConsultMinute: Number(option.key) }));
                                }}
                                styles={{ dropdown: { width: 70 } }}
                                disabled={isReadOnly}
                              />
                            </Stack>
                          </div>
                        </Stack.Item>
                        <Stack.Item grow>
                          <div className={styles.modalFormField}>
                            <TextField
                              id="modal-edc"
                              label="EDC"
                              type="date"
                              tabIndex={consultTabOrder.edc}
                              value={formData.EDC}
                              onChange={(_, newValue): void =>
                                setFormData(prev => ({ ...prev, EDC: newValue ?? '' }))
                              }
                              styles={{ root: { width: '100%' } }}
                              disabled={isReadOnly}
                              readOnly={isReadOnly}
                            />
                          </div>
                        </Stack.Item>
                      </Stack>

                      <Stack horizontal tokens={{ childrenGap: 16 }} className={styles.consultFormRow}>
                        <Stack.Item grow>
                          <div className={styles.modalFormField}>
                            <TextField
                              id="modal-title"
                              label="Last Name"
                              required
                              tabIndex={consultTabOrder.lastName}
                              value={formData.Title}
                              errorMessage={lastNameError}
                              onChange={(_, newValue): void => {
                                setLastNameError(undefined);
                                setFormData(prev => ({ ...prev, Title: newValue ?? '' }));
                              }}
                              disabled={isReadOnly}
                              readOnly={isReadOnly}
                              styles={{ root: { width: '100%' } }}
                            />
                          </div>
                        </Stack.Item>
                        <Stack.Item grow>
                          <div className={styles.modalFormField}>
                            <TextField
                              id="modal-firstname"
                              label="First Name"
                              tabIndex={consultTabOrder.firstName}
                              value={formData.FirstName}
                              onChange={(_, newValue): void =>
                                setFormData(prev => ({ ...prev, FirstName: newValue ?? '' }))
                              }
                              disabled={isReadOnly}
                              readOnly={isReadOnly}
                              styles={{ root: { width: '100%' } }}
                            />
                          </div>
                        </Stack.Item>
                      </Stack>

                      <Stack horizontal tokens={{ childrenGap: 16 }} className={styles.consultFormRow}>
                        <Stack.Item grow>
                          <div className={styles.modalFormField}>
                            <TextField
                              id="modal-maternaldob"
                              label="Maternal DOB"
                              type="date"
                              tabIndex={consultTabOrder.maternalDob}
                              value={formData.MaternalDOB}
                              onChange={(_, newValue): void =>
                                setFormData(prev => ({ ...prev, MaternalDOB: newValue ?? '' }))
                              }
                              disabled={isReadOnly}
                              readOnly={isReadOnly}
                            />
                          </div>
                        </Stack.Item>
                        <Stack.Item grow>
                          <div className={styles.modalFormField}>
                            <TextField
                              id="modal-maternalmrn"
                              label="Maternal MRN"
                              tabIndex={consultTabOrder.maternalMrn}
                              value={formData.MaternalMRN}
                              onChange={(_, newValue): void =>
                                setFormData(prev => ({ ...prev, MaternalMRN: newValue ?? '' }))
                              }
                              disabled={isReadOnly}
                              readOnly={isReadOnly}
                            />
                          </div>
                        </Stack.Item>
                      </Stack>

                      <div className={`${styles.modalFormField} ${(styles as Record<string, string>).modalFormSectionStart}`}>
                        <Label>Additional Consult Information</Label>
                        <RichTextEditor
                          value={formData.AdditionalConsulInfo ?? ''}
                          tabIndex={consultTabOrder.additionalConsultInfo}
                          onChange={(value: string): void => {
                            setFormData(prev => ({ ...prev, AdditionalConsulInfo: value }));
                          }}
                          readOnly={isReadOnly}
                        />
                      </div>

                      <Stack horizontal tokens={{ childrenGap: 16 }} className={`${styles.consultFormRow} ${(styles as Record<string, string>).modalFormSectionStart}`}>
                        <Stack.Item grow>
                          <div className={styles.modalFormField}>
                            <Label htmlFor="modal-pregnancy">Pregnancy Status</Label>
                            <Dropdown
                              id="modal-pregnancy"
                              tabIndex={consultTabOrder.pregnancyStatus}
                              selectedKey={formData.PregnancyStatus || ''}
                              options={[
                                { key: '', text: 'Please Select...' },
                                ...getSanitizedChoices('PregnancyStatus').map(opt => ({
                                  key: opt,
                                  text: opt
                                }))
                              ]}
                              onChange={(_, option): void => {
                                if (!option) return;
                                setFormData(prev => ({ ...prev, PregnancyStatus: String(option.key) }));
                              }}
                              ariaLabel="Pregnancy Status"
                              disabled={isReadOnly}
                            />
                          </div>
                        </Stack.Item>
                        <Stack.Item grow>
                          <div className={`${styles.modalFormField} ${(styles as Record<string, string>).modalFormAttachments}`}>
                            <Label className={(styles as Record<string, string>).modalFormAttachmentsLabel}>Attachments</Label>
                            {consultExistingAttachments.length === 0 && consultPendingFiles.length === 0 && (
                              <p className={(styles as Record<string, string>).modalFormAttachmentsEmpty}>There is nothing attached.</p>
                            )}
                            {consultExistingAttachments.length > 0 && (
                              <ul className={(styles as Record<string, string>).modalFormAttachmentsList}>
                                {consultExistingAttachments.map(att => (
                                  <li key={att.url}>
                                    <a href={att.url} target="_blank" rel="noreferrer">{att.name}</a>
                                    {modalMode === 'edit' && modalItemId !== null && (
                                      <DefaultButton
                                        text="Remove"
                                        iconProps={{ iconName: 'Delete' }}
                                        styles={{ root: { marginLeft: 8, minWidth: 0, padding: '0 6px' } }}
                                        onClick={(e): void => {
                                          e.preventDefault();
                                          deleteAttachment(modalItemId, att.url).catch(() => {
                                            // ignore; attachment list will refresh on next open
                                          });
                                        }}
                                      />
                                    )}
                                  </li>
                                ))}
                              </ul>
                            )}
                            {consultPendingFiles.length > 0 && (
                              <Stack tokens={{ childrenGap: 4 }}>
                                <Text variant="small">
                                  Uploaded files: {consultPendingFiles.map(f => f.name).join(', ')}
                                </Text>
                                {(modalMode === 'new' || modalMode === 'edit') && (
                                  <DefaultButton
                                    text="Clear Uploads"
                                    iconProps={{ iconName: 'Cancel' }}
                                    styles={{ root: { alignSelf: 'flex-start' } }}
                                    onClick={(): void => { setConsultPendingFiles([]); consultPendingFilesRef.current = []; }}
                                  />
                                )}
                              </Stack>
                            )}
                            {(modalMode === 'new' || modalMode === 'edit') && (
                              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                                <DefaultButton
                                  iconProps={{ iconName: 'Upload' }}
                                  text="Choose file"
                                  tabIndex={consultTabOrder.chooseFile}
                                  className={(styles as Record<string, string>).modalFormAttachLink}
                                  onClick={(): void => consultFileInputRef.current?.click()}
                                  aria-label="Choose file"
                                  disabled={formLoading}
                                />
                                <Text variant="small" styles={{ root: { opacity: 0.9 } }}>
                                  {consultPendingFiles.length > 0 ? consultPendingFiles.map(f => f.name).join(', ') : 'No file chosen'}
                                </Text>
                              </Stack>
                            )}
                            <input
                              type="file"
                              multiple
                              ref={consultFileInputRef}
                              style={{ display: 'none' }}
                              onChange={(e): void => {
                                const files = Array.from(e.target.files ?? []);
                                if (files.length === 0) return;
                                setConsultPendingFiles(prev => {
                                  const next = [...prev, ...files];
                                  consultPendingFilesRef.current = next;
                                  return next;
                                });
                                e.target.value = '';
                              }}
                            />
                          </div>
                        </Stack.Item>
                      </Stack>

                      {formError && <p className={styles.modalFormError}>{formError}</p>}
                    </Stack>
                  </div>
                  );
                })()}
                {!formLoading && activeFormTab === 'medical' && (() => {
                  const isReadOnly = modalMode === 'view';
                  return (
                    <>
                      <div className={styles.modalFormField}>
                        <Label>Fetal Diagnoses</Label>
                        <RichTextEditor
                          value={formData.FetalDiagnoses ?? ''}
                          tabIndex={medicalTabOrder.fetalDiagnoses}
                          onChange={(value: string): void => {
                            setFormData(prev => ({ ...prev, FetalDiagnoses: value }));
                          }}
                          readOnly={isReadOnly}
                        />
                      </div>
                      <div className={(styles as Record<string, string>).modalFormTwoCol}>
                        <div className={(styles as Record<string, string>).modalFormCol}>
                          <div className={styles.modalFormField}>
                            <label htmlFor="modal-primaryorgan">Primary Organ System</label>
                            <select
                              id="modal-primaryorgan"
                              tabIndex={medicalTabOrder.primaryOrgan}
                              value={formData.PrimaryOrganSystem}
                              onChange={(e): void => setFormData(prev => ({ ...prev, PrimaryOrganSystem: e.target.value }))}
                              disabled={isReadOnly}
                              className={(styles as Record<string, string>).modalFormSelect}
                            >
                              <option value="">Please Select...</option>
                              {getSanitizedChoices('PrimaryOrganSystem').map((opt: string) => (
                                <option key={opt} value={opt}>{opt}</option>
                              ))}
                            </select>
                          </div>
                        </div>
                        <div className={(styles as Record<string, string>).modalFormCol}>
                          <div className={styles.modalFormField}>
                            <label htmlFor="modal-secondaryorgan">Secondary Organ System</label>
                            <select
                              id="modal-secondaryorgan"
                              tabIndex={medicalTabOrder.secondaryOrgan}
                              value={formData.SecondaryOrganSystem}
                              onChange={(e): void => setFormData(prev => ({ ...prev, SecondaryOrganSystem: e.target.value }))}
                              disabled={isReadOnly}
                              className={(styles as Record<string, string>).modalFormSelect}
                            >
                              <option value="">Please Select...</option>
                              {getSanitizedChoices('SecondaryOrganSystem').map((opt: string) => (
                                <option key={opt} value={opt}>{opt}</option>
                              ))}
                            </select>
                          </div>
                        </div>
                      </div>
                      <div className={styles.modalFormField}>
                        <Label>Prenatal Imaging</Label>
                        <RichTextEditor
                          value={formData.PrenatalImaging ?? ''}
                          tabIndex={medicalTabOrder.prenatalImaging}
                          onChange={(value: string): void => {
                            setFormData(prev => ({ ...prev, PrenatalImaging: value }));
                          }}
                          readOnly={isReadOnly}
                        />
                      </div>
                      <div className={styles.modalFormField}>
                        <Label>Genetic Testing</Label>
                        <RichTextEditor
                          value={formData.GeneticTesting ?? ''}
                          tabIndex={medicalTabOrder.geneticTesting}
                          onChange={(value: string): void => {
                            setFormData(prev => ({ ...prev, GeneticTesting: value }));
                          }}
                          readOnly={isReadOnly}
                        />
                      </div>
                      <div className={styles.modalFormField}>
                        <label htmlFor="modal-scheduled-med">Scheduled Date for Delivery</label>
                        <input
                          id="modal-scheduled-med"
                          type="date"
                          tabIndex={medicalTabOrder.scheduledDelivery}
                          value={formData.ScheduledDateDelivery}
                          onChange={(e): void => setFormData(prev => ({ ...prev, ScheduledDateDelivery: e.target.value }))}
                          disabled={isReadOnly}
                          readOnly={isReadOnly}
                        />
                      </div>
                      <div className={(styles as Record<string, string>).modalFormTwoCol}>
                        <div className={(styles as Record<string, string>).modalFormCol}>
                          <div className={styles.modalFormField}>
                            <label htmlFor="modal-deliverymode">Planned Mode of Delivery</label>
                            <select
                              id="modal-deliverymode"
                              tabIndex={medicalTabOrder.deliveryMode}
                              value={formData.DeliveryMode}
                              onChange={(e): void => setFormData(prev => ({ ...prev, DeliveryMode: e.target.value }))}
                              disabled={isReadOnly}
                              className={(styles as Record<string, string>).modalFormSelect}
                            >
                              <option value="">Please Select...</option>
                              {getSanitizedChoices('DeliveryMode').map((opt: string) => (
                                <option key={opt} value={opt}>{opt}</option>
                              ))}
                            </select>
                          </div>
                        </div>
                        <div className={(styles as Record<string, string>).modalFormCol}>
                          <div className={styles.modalFormField}>
                            <label htmlFor="modal-deliveryloc">Planned Delivery Location</label>
                            <select
                              id="modal-deliveryloc"
                              tabIndex={medicalTabOrder.deliveryLocation}
                              value={formData.DeliveryLocation}
                              onChange={(e): void => setFormData(prev => ({ ...prev, DeliveryLocation: e.target.value }))}
                              disabled={isReadOnly}
                              className={(styles as Record<string, string>).modalFormSelect}
                            >
                              <option value="">Please Select...</option>
                              {getSanitizedChoices('DeliveryLocation').map((opt: string) => (
                                <option key={opt} value={opt}>{opt}</option>
                              ))}
                            </select>
                          </div>
                        </div>
                      </div>
                      <div className={(styles as Record<string, string>).modalFormTwoCol}>
                        <div className={(styles as Record<string, string>).modalFormCol}>
                          <div className={styles.modalFormField}>
                            <label htmlFor="modal-med-baby-mrn">Baby MRN</label>
                            <input
                              id="modal-med-baby-mrn"
                              type="text"
                              tabIndex={medicalTabOrder.babyMrn}
                              value={formData.BabyMRN}
                              onChange={(e): void => setFormData(prev => ({ ...prev, BabyMRN: e.target.value }))}
                              disabled={isReadOnly}
                              readOnly={isReadOnly}
                            />
                          </div>
                          <div className={styles.modalFormField}>
                            <label htmlFor="modal-med-baby3-mrn">Baby 3 MRN</label>
                            <input
                              id="modal-med-baby3-mrn"
                              type="text"
                              tabIndex={medicalTabOrder.baby3Mrn}
                              value={formData.Baby3MRN}
                              onChange={(e): void => setFormData(prev => ({ ...prev, Baby3MRN: e.target.value }))}
                              disabled={isReadOnly}
                              readOnly={isReadOnly}
                            />
                          </div>
                        </div>
                        <div className={(styles as Record<string, string>).modalFormCol}>
                          <div className={styles.modalFormField}>
                            <label htmlFor="modal-med-baby2-mrn">Baby 2 MRN</label>
                            <input
                              id="modal-med-baby2-mrn"
                              type="text"
                              tabIndex={medicalTabOrder.baby2Mrn}
                              value={formData.Baby2MRN}
                              onChange={(e): void => setFormData(prev => ({ ...prev, Baby2MRN: e.target.value }))}
                              disabled={isReadOnly}
                              readOnly={isReadOnly}
                            />
                          </div>
                          <div className={styles.modalFormField}>
                            <label htmlFor="modal-med-baby4-mrn">Baby 4 MRN</label>
                            <input
                              id="modal-med-baby4-mrn"
                              type="text"
                              tabIndex={medicalTabOrder.baby4Mrn}
                              value={formData.Baby4MRN}
                              onChange={(e): void => setFormData(prev => ({ ...prev, Baby4MRN: e.target.value }))}
                              disabled={isReadOnly}
                              readOnly={isReadOnly}
                            />
                          </div>
                        </div>
                      </div>
                      <div className={styles.modalFormField}>
                        <Label>Delivery Room Considerations</Label>
                        <RichTextEditor
                          value={formData.DeliveryRoomConsid ?? ''}
                          tabIndex={medicalTabOrder.deliveryRoom}
                          onChange={(value: string): void => {
                            setFormData(prev => ({ ...prev, DeliveryRoomConsid: value }));
                          }}
                          readOnly={isReadOnly}
                        />
                      </div>
                      <div className={styles.modalFormField}>
                        <label htmlFor="modal-pge">PGEs Dependent</label>
                        <select
                          id="modal-pge"
                          tabIndex={medicalTabOrder.pgeDependent}
                          value={formData.PGEDependent}
                          onChange={(e): void => setFormData(prev => ({ ...prev, PGEDependent: e.target.value }))}
                          disabled={isReadOnly}
                          className={(styles as Record<string, string>).modalFormSelect}
                        >
                          <option value="">Please Select...</option>
                          {getSanitizedChoices('PGEDependent').map((opt: string) => (
                            <option key={opt} value={opt}>{opt}</option>
                          ))}
                        </select>
                      </div>
                      <div className={styles.modalFormField}>
                        <Label>Postnatal Care Plans</Label>
                        <RichTextEditor
                          value={formData.PostnatalCarePlans ?? ''}
                          tabIndex={medicalTabOrder.postnatalCare}
                          onChange={(value: string): void => {
                            setFormData(prev => ({ ...prev, PostnatalCarePlans: value }));
                          }}
                          readOnly={isReadOnly}
                        />
                      </div>
                      <div className={styles.modalFormField}>
                        <Label>Family Social Considerations</Label>
                        <RichTextEditor
                          value={formData.FamilySocialConsiderations ?? ''}
                          tabIndex={medicalTabOrder.familySocial}
                          onChange={(value: string): void => {
                            setFormData(prev => ({ ...prev, FamilySocialConsiderations: value }));
                          }}
                          readOnly={isReadOnly}
                        />
                      </div>
                      <div className={styles.modalFormField}>
                        <Label>Active Medical Concerns/Pending Issues</Label>
                        <RichTextEditor
                          value={formData.ActiveConcernsPendingIssues ?? ''}
                          tabIndex={medicalTabOrder.activeConcerns}
                          onChange={(value: string): void => {
                            setFormData(prev => ({ ...prev, ActiveConcernsPendingIssues: value }));
                          }}
                          readOnly={isReadOnly}
                        />
                      </div>
                    </>
                  );
                })()}
                {!formLoading && activeFormTab === 'administrative' && (() => {
                  const isReadOnly = modalMode === 'view';
                  return (
                    <div className={(styles as Record<string, string>).modalFormTwoCol}>
                      <div className={(styles as Record<string, string>).modalFormCol}>
                        <div className={styles.modalFormField}>
                          <label htmlFor="modal-prenatal-team">Prenatal Team Discussion (Date)</label>
                          <input
                            id="modal-prenatal-team"
                            type="date"
                            tabIndex={administrativeTabOrder.prenatalTeam}
                            value={formData.PrenatalTeamDiscussion}
                            onChange={(e): void => setFormData(prev => ({ ...prev, PrenatalTeamDiscussion: e.target.value }))}
                            disabled={isReadOnly}
                            readOnly={isReadOnly}
                          />
                        </div>
                        <div className={styles.modalFormField}>
                          <label htmlFor="modal-baby-dob">Baby DOB</label>
                          <input
                            id="modal-baby-dob"
                            type="date"
                            tabIndex={administrativeTabOrder.babyDob}
                            value={formData.BabyDOB}
                            onChange={(e): void => setFormData(prev => ({ ...prev, BabyDOB: e.target.value }))}
                            disabled={isReadOnly}
                            readOnly={isReadOnly}
                          />
                        </div>
                        <div className={styles.modalFormField}>
                          <label htmlFor="modal-baby2-mrn">Baby 2 MRN</label>
                          <input
                            id="modal-baby2-mrn"
                            type="text"
                            tabIndex={administrativeTabOrder.baby2Mrn}
                            value={formData.Baby2MRN}
                            title="Edit this field in the Medical tab"
                            disabled
                            readOnly
                          />
                        </div>
                        <div className={styles.modalFormField}>
                          <label htmlFor="modal-baby4-mrn">Baby 4 MRN</label>
                          <input
                            id="modal-baby4-mrn"
                            type="text"
                            tabIndex={administrativeTabOrder.baby4Mrn}
                            value={formData.Baby4MRN}
                            title="Edit this field in the Medical tab"
                            disabled
                            readOnly
                          />
                        </div>
                      </div>
                      <div className={(styles as Record<string, string>).modalFormCol}>
                        <div className={styles.modalFormField}>
                          <label htmlFor="modal-multi-team">Multidisciplinary Team Discussion (date)</label>
                          <input
                            id="modal-multi-team"
                            type="date"
                            tabIndex={administrativeTabOrder.multiTeam}
                            value={formData.MultiTeamDiscussion}
                            onChange={(e): void => setFormData(prev => ({ ...prev, MultiTeamDiscussion: e.target.value }))}
                            disabled={isReadOnly}
                            readOnly={isReadOnly}
                          />
                        </div>
                        <div className={styles.modalFormField}>
                          <label htmlFor="modal-baby-mrn">Baby MRN</label>
                          <input
                            id="modal-baby-mrn"
                            type="text"
                            tabIndex={administrativeTabOrder.babyMrn}
                            value={formData.BabyMRN}
                            title="Edit this field in the Medical tab"
                            disabled
                            readOnly
                          />
                        </div>
                        <div className={styles.modalFormField}>
                          <label htmlFor="modal-baby3-mrn">Baby 3 MRN</label>
                          <input
                            id="modal-baby3-mrn"
                            type="text"
                            tabIndex={administrativeTabOrder.baby3Mrn}
                            value={formData.Baby3MRN}
                            title="Edit this field in the Medical tab"
                            disabled
                            readOnly
                          />
                        </div>
                      </div>
                    </div>
                  );
                })()}
              </div>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
