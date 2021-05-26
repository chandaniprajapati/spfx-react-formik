import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReactFormikProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
}
