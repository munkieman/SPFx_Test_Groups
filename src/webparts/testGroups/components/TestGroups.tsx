import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './TestGroups.module.scss';
import type { ITestGroupsProps } from './ITestGroupsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';

const TestGroups: React.FC<ITestGroupsProps> = (props) => {
  const {
    description,
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName,
    context
  } = props;

  const [groupMembers, setGroupMembers] = useState<any[]>([]);

  useEffect(() => {
    const fetchGroupMembers = async () => {
      try {
        const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient();
        const response = await client.api('/groups/68d9eb2c-06f7-40ed-bd99-a5a35fab0275/members').get();

        if (response && response.value) {
          setGroupMembers(response.value);
        } else {
          console.warn('No group members found.');
        }
      } catch (error) {
        console.error('Error fetching group members:', error);
      }
    };

    fetchGroupMembers();
  }, [context]);

  return (
    <section className={`${styles.testGroups} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <p>
          The SharePoint Framework (SPFx) is an extensibility model for Microsoft Viva, Microsoft Teams, and SharePoint.
          It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign-On, automatic hosting, and industry-standard tooling.
        </p>
        <h4>Learn more about SPFx development:</h4>
        <ul className={styles.links}>
          <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
          <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
          <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
        </ul>
      </div>
      <div>
        <h4>Group Members:</h4>
        {groupMembers.length > 0 ? (
          <ul>
            {groupMembers.map((member, index) => (
              <li key={index}>
                {member.displayName} ({member.mail || 'No email available'})
              </li>
            ))}
          </ul>
        ) : (
          <p>No members found in this group.</p>
        )}
      </div>
    </section>
  );
};

export default TestGroups;
