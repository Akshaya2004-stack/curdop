import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICurdOperationsSpfxProps {
  description: string;
  listName: string;
  // isDarkTheme: boolean;
  context: WebPartContext;
}
