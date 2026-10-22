import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IBsupDocGen1Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:WebPartContext;
  adminEmailWbProb:string;  
  managerEmailWbProb:string;
}
