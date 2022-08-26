export interface IPersonalGreetingProps {
  currentDate: string;
  message: string;
  optionalMessage: string;
  // photoSize: string;  
  // photoUrl: string;
  primaryTextColor: string;
  secondaryTextColor: string;
  tertiaryTextColor: string;
  primaryTextSize: number;
  secondaryTextSize: number;
  tertiaryTextSize: number;  
  position: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
