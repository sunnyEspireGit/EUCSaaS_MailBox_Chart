export interface IEmailReportProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  
  applyDateValue:(startDate,endDate)=>void;

  onStartDateChanged: (date: Date | null | undefined) => void;
  onEndDateChanged: (date: Date | null | undefined) => void;
}
