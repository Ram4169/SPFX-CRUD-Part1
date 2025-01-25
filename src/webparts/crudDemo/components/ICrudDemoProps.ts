import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICrudDemoProps {
  description: string;
  context: WebPartContext;
  selectedLibrary: string;
}
