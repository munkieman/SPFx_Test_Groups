import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './TestGraphClient.module.scss';
import type { ITestGraphClientProps } from './ITestGraphClientProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";

//interface IMember {
//  id: string;
//  displayName: string;
//}

interface ITag {
  id: string;
  displayName: string;
}

interface DialogProps {
  type: "error" | "success" | "warning" | "info";
  message: string | null;
  onClose: () => void;
}

//client secret value : FoA8Q~MyIfYbTcjmarpbpOjb07VBKKksYcIYwaiA
//client secret id    : 0af0eb1b-f72c-495f-b3bd-f9273c7edf6d

const TestGraphClient: React.FC<ITestGraphClientProps> = (props) => {
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
  const [tags, setTags] = useState<ITag[]>([]);
  const [dialogMessage, setDialogMessage] = useState<string | null>(null);
  const [isDialogOpen, setIsDialogOpen] = useState<boolean>(false);
  const [dialogType, setDialogType] = useState<"error" | "success" | "warning" | "info">("info"); // Default type

  // Max Prod Team
  //const teamName = "ExpensesChat";
  const teamID = "68d9eb2c-06f7-40ed-bd99-a5a35fab0275";

  //const teamName = "Teams Testing";
  //const teamID = "a3cce0fc-52f7-4928-8f2b-14102e5ad6ca";


  // Max Dev Team
  //const teamName = "TestChat";
  //const teamID = "696dfe67-e76f-4bf8-8ab6-8abfcb16552e";

  //const channelName = "General";
  const userEmail = props.context.pageContext.user.email;

  //https://teams.microsoft.com/l/channel/19%3AWELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1%40thread.tacv2/General?groupId=68d9eb2c-06f7-40ed-bd99-a5a35fab0275&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb

  //https://teams.microsoft.com/l/channel/19%3Aec62a56976504a9da063458459e73b34%40thread.tacv2/General?groupId=f5de9aad-7f98-498d-a5a0-b1a59254265c&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb

  const Dialog: React.FC<DialogProps> = ({ type, message, onClose }) => {
    if (!message) return null;
  
    return (
      <div className={styles.dialogOverlay}>
        <div className={`${styles.dialogBox} ${styles[type]}`}>
          <p>{message}</p>
          <button onClick={onClose}>Close</button>
        </div>
      </div>
    );
  };

  const showDialog = (type: "error" | "success" | "warning" | "info", message: string) => {
    setDialogType(type);
    setDialogMessage(message);
    setIsDialogOpen(true);
  };

  const getTeamTags = (): void => {  
    console.log("Fetching tags for team ID:", teamID);

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

  const addMember = async (): Promise<void> => {
    /*
    try {      
      const client = await context.msGraphClientFactory.getClient('3');
      //const userResponse = await client.api(`/users/${userEmail}`).version('v1.0').get();
      //const userId = userResponse.id;
      const user = await client.api('/me').get();
      const userId = user.id;
      const today = new Date().toISOString();
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
        
        const conversationMember = {
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${userId}`,
            visibleHistoryStartDateTime: `${today}`,
            roles: ['owner']
        };

        await client.api('/chats/19:3AWELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1@thread.v2/members')
          .post(conversationMember);    
        
        const addUserResponse = await client.api(`/groups/${teamID}/members/$ref`)
        .post({
            "@odata.id": `https://graph.microsoft.com/v1.0/users/${userId}`
        });
        
            //"@odata.type": "#microsoft.graph.aadUserConversationMember",
            //"roles": ["owner"],
            //"isHistoryIncluded": false,
            //"visibleHistoryStartDateTime": `${today}`  //"visibleHistoryStartDateTime": "2025-04-10T14:58:56.284Z"

        // Add user to the team
        
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

        if (!addUserResponse.ok) {
          const errorText = await addUserResponse.text();
          throw new Error(`Failed to add user to the chat: ${errorText}`);
        }

        console.log("User successfully added to the chat");
        
      } else {
        console.log("User is already a member of the chat channel");
        //getTeamTags();
      }                
      //getTeamTags();

    } catch (error: any) {
      console.error("Error in checkMember:", error.message);
    }
    */

    const client = await context.aadHttpClientFactory.getClient("https://graph.microsoft.com");

    //Fetch the user ID
    const userResponse = await client.get(
      `https://graph.microsoft.com/v1.0/users/${userEmail}`,
      AadHttpClient.configurations.v1
    );
    const userData = await userResponse.json();
    const userId = userData.id;      

    let userIsMember = false;
    let membersUrl = `https://graph.microsoft.com/v1.0/teams/${teamID}/members`;

    console.log("userID",userId,userData);

    do {
      // Fetch team members
      const membersResponse = await client.get(membersUrl, AadHttpClient.configurations.v1);
      if (!membersResponse.ok) {
        throw new Error("Failed to fetch members");
      }
      const membersData = await membersResponse.json();
      
      console.log("Members Data:", membersData);  // Log member data for debugging
      console.log("Comparing userId:", userId);  // Log userId for debugging

      // Check if the user is a member in this page of members
      userIsMember = membersData.value.some((m: any) => m.id === userId);
  
      // If the user is found, exit the loop
      if (userIsMember) {
        console.log("User is a member of the team!");
        break;
      }
  
      // If pagination exists, update the membersUrl to the next page
      membersUrl = membersData["@odata.nextLink"];
      
    } while (membersUrl && !userIsMember); 

    if (!userIsMember) {
      showDialog("info","Adding you to the chat channel. Please wait...");    
      console.log("User is not a member, adding to the chat channel...");

      // Add user to the team
      const addUserResponse: HttpClientResponse = await client.post(
        `https://graph.microsoft.com/v1.0/teams/${teamID}/members`, //channels/${channel.id}/members`,
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
            //"isHistoryIncluded": false, 
            //"visibleHistoryStartDateTime": null
          })
        }
      );

      if (!addUserResponse.ok) {
        const errorText = await addUserResponse.text();
        showDialog("error","Failed to add user to the chat: " + errorText);
        throw new Error(`Failed to add user to the chat: ${errorText}`);
      }

      showDialog("success","You have successfully joined the chat!");

      //setIsChatDisabled(true);
      console.log("User successfully added to the chat");
    }else{
      showDialog("success","You are already a member of this chat channel.");        
      console.log("User is already a member of the chat channel");          
    }      
  };  

  useEffect(() => {

    const fetchGroupMembers = async ():Promise<void> => {
      try {
        const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
        const response = await client.api(`/groups/${teamID}/members`).get();

        if (response && response.value) {
          setGroupMembers(response.value);
          getTeamTags();
        } else {
          console.warn('No group members found.');
        }
      } catch (error) {
        console.error('Error fetching group members:', error);
      }
    };

    fetchGroupMembers();
    //if (groupMembers.length > 0) {
    //}

  }, [context]);

  return (
    <section className={`${styles.testGraphClient} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <div>
          {isDialogOpen && <Dialog type={dialogType} message={dialogMessage} onClose={() => setIsDialogOpen(false)} />}
        </div>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
        <div>User Email: {escape(userEmail)}</div>
      </div>
      <button onClick={addMember}>Join Chat</button>
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
      <div>
        <h4>Tags</h4>
        {tags.length > 0 ? (
          <ul>
            {tags.map((tag, index) => (
              <li key={index}>
                {tag.displayName} : ({tag.id})
              </li>
            ))}
          </ul>
        ) : (
          <p>No tags found in this team.</p>
        )}
      </div>           
    </section>
  );
};

export default TestGraphClient;
