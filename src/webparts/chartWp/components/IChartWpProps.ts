
import { DynamicProperty } from '@microsoft/sp-component-base';

export interface IChartWpProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  // startDate: DynamicProperty<String>;
  startDate: DynamicProperty<String>;
  endDate: DynamicProperty<String>;
  dateRange: DynamicProperty<String>;
}
