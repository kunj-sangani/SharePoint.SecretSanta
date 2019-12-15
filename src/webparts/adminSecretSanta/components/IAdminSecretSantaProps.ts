import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAdminSecretSantaProps {
  description: string;
  adminUserEmail:string;
  context:WebPartContext;
  lists: string;
}
