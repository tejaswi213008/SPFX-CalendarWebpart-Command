import type { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICalendarviewforSpxProps {
  context: WebPartContext;
  listTitle: string;
  /** Optional. When set, load list from this site URL. Leave blank to use current site. */
  siteUrl?: string;
  /** Optional. Internal name for NewForm/EditForm/ViewForm URLs (e.g. SPFXRMNHCalendar). When set, form links use this instead of list title. */
  listInternalName?: string;
  /** Optional. Power Apps form URL; when set, event click opens modal instead of EditForm. */
  powerAppsFormUrl?: string;
  /** Optional. Append to Power Apps URL to open in Edit mode (e.g. "Mode=Edit" or "FormMode=Edit"). Use whatever your app expects. */
  powerAppsEditModeParam?: string;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
