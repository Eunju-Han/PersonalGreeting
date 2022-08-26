import * as React from 'react';
import styles from './PersonalGreeting.module.scss';
import { IPersonalGreetingProps } from './IPersonalGreetingProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PersonalGreeting extends React.Component<IPersonalGreetingProps, {}> {

  public render(): React.ReactElement<IPersonalGreetingProps> {
    const {
      currentDate,
      message,
      optionalMessage,      
      primaryTextColor,
      secondaryTextColor,
      tertiaryTextColor,
      primaryTextSize,
      secondaryTextSize,
      tertiaryTextSize,  
      position,      
      // isDarkTheme,
      // environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const GreetingMessageCustStyles: React.CSSProperties = { 'color': primaryTextColor,'font-size': primaryTextSize,'text-align': position } as React.CSSProperties;
    const OptionalMessageCustStyles: React.CSSProperties = { 'color': secondaryTextColor, 'font-size': secondaryTextSize, 'text-align': position } as React.CSSProperties;
    const dateCustStyles: React.CSSProperties = { 'color': tertiaryTextColor, 'font-size': tertiaryTextSize, 'text-align': position } as React.CSSProperties;

    // console.log("*** userDisplayName: " + userDisplayName);
    return (
      <section className={`${styles.personalGreeting} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={ styles.customClass }>
        {/* <div className={ styles.welcome }> */}
          <p style={GreetingMessageCustStyles}>{`${escape(message)}${userDisplayName}`}</p>
          {optionalMessage ? <p style={OptionalMessageCustStyles}>
          {`${escape(optionalMessage)}`}</p>: ''}
          <p style={dateCustStyles}>{currentDate}</p>
        </div>
      </section>
    );
  }
}
