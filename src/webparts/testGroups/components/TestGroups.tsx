import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './TestGroups.module.scss';
import type { ITestGroupsProps } from './ITestGroupsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';

//interface IMember {
//  id: string;
//  displayName: string;
//}

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
  //const [members, setMembers] = useState<IMember[]>([]);
  //const [loading, setLoading] = useState<boolean>(false);
  //const [error, setError] = useState<string | null>(null);
  //const [tags, setTags] = useState<ITag[]>([]);

  //const teamName = "ExpensesChat";
  const teamID = "68d9eb2c-06f7-40ed-bd99-a5a35fab0275";

  //const teamName = "Teams Testing";
  //const teamID = "a3cce0fc-52f7-4928-8f2b-14102e5ad6ca";

  //const channelName = "General";
  const userEmail = props.context.pageContext.user.email;

  //https://teams.microsoft.com/l/channel/19%3AWELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1%40thread.tacv2/General?groupId=68d9eb2c-06f7-40ed-bd99-a5a35fab0275&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb

  useEffect(() => {

    const checkMember = async (): Promise<void> => {
      try {
        const client = await context.msGraphClientFactory.getClient('3');
        //const userResponse = await client.api(`/users/${userEmail}`).version('v1.0').get();
        //const userId = userResponse.id;
        const user = await client.api('/me').get();
        const userId = user.id;
        const today = new Date().toISOString;
        const userIsMember = groupMembers.some(member => member.id === userId);
        
        console.log("Is user a member of the team:", userIsMember);
        console.log("User ID:", userId);
        console.log("groupmembers",groupMembers);
        console.log("today",today);

        if (!userIsMember) {
          console.log("User is not a member, adding to the chat channel...");
    
          // Add user to the team
          
          /* this works but just adds the user to the team 

          const directoryObject = {            
            '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
          };
          
          await client.api(`/groups/${teamID}/members/$ref`)
            .post(directoryObject);          
         
          */

          const conversationMember = {
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${userId}`,
              visibleHistoryStartDateTime: `${today}`,
              roles: ['owner']
          };

          await client.api('/chats/19:3AWELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1@thread.v2/members')
            .post(conversationMember);

          /*
            const addUserResponse = await client.api(`/groups/${teamID}/members/$ref`)
          .post({
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              "@odata.type": "#microsoft.graph.aadUserConversationMember",
              "roles": ["member"],
              "@odata.id": `https://graph.microsoft.com/v1.0/users/${userId}`,
              "isHistoryIncluded": false,
              "visibleHistoryStartDateTime": null
            })
          });
          */

          // Add user to the team
          /*
          const addUserResponse: HttpClientResponse = await client.post(
            `https://graph.microsoft.com/v1.0/teams/${team.id}/members`, //channels/${channel.id}/members`,
             AadHttpClient.configurations.v1,
            {
              headers: { 
                "Content-Type": "application/json"
                //Authorization : `Bearer ${accessToken}`,
              },
              body: JSON.stringify({
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["member"],
                "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userId}`,
                "isHistoryIncluded": false, 
                "visibleHistoryStartDateTime": null
              })
            }
          );
          */
  
          //if (!addUserResponse.ok) {
          //  const errorText = await addUserResponse.text();
          //  throw new Error(`Failed to add user to the chat: ${errorText}`);
          //}

          console.log("User successfully added to the chat");
          
        } else {
          console.log("User is already a member of the chat channel");
        }                
        

      } catch (error: any) {
        console.error("Error in checkMember:", error.message);
      }
    };

    const fetchGroupMembers = async ():Promise<void> => {
      try {
        const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
        const response = await client.api(`/groups/${teamID}/members`).get();

        if (response && response.value) {
          setGroupMembers(response.value);
        } else {
          console.warn('No group members found.');
        }
      } catch (error) {
        console.error('Error fetching group members:', error);
      }
      checkMember();
    };

/*

    const fetchChannelMembers = async ():Promise<void> => {
      try {
        //setLoading(true);
  
        // Get Microsoft Graph API client
        const client : MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
  
        // Get joined teams
        const teamsResponse = await client.api('/me/joinedTeams')
          .version('v1.0')
          .get();
        
        if (!teamsResponse) throw new Error("Failed to fetch teams");
  
      } catch (err: any) {
        console.log("Error fetching channel members:", err);
        //setError(err.message);
      } finally {
        //setLoading(false);
      }
    };
    const getTeamTags = (): void => {  
      context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client
            .api(`/teams/${teamID}/tags`)
            .version('v1.0')
            .get((error, response: any) => {
              if (error) {
                console.error('Error fetching tags:', error);
                return;
              }
              setTags(response.value);
            });
        });
    };    
*/

    fetchGroupMembers();
    //if (groupMembers.length > 0) {
    //}

    //getTeamTags();

  }, [context]);

  return (
    <section className={`${styles.testGroups} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
        <div>User Email: {escape(userEmail)}</div>
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
