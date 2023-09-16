import { WebPartContext } from "@microsoft/sp-webpart-base"; //import WebPartContext here

export interface IMyMeetingsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  websiteUrl: string; // include this
  spcontext: WebPartContext; // include spcontext here and make changes to MyMeetingsWebPart.ts file
}
