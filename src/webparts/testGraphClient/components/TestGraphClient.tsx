import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './TestGraphClient.module.scss';
import type { ITestGraphClientProps } from './ITestGraphClientProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { AadHttpClient } from "@microsoft/sp-http";

//interface IMember {
//  id: string;
//  displayName: string;
//}


interface ITag {
  id: string;
  displayName: string;
}

/*
interface DialogProps {
  type: "error" | "success" | "warning" | "info";
  message: string | null;
  onClose: () => void;
}
*/

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
  //const [dialogMessage, setDialogMessage] = useState<string | null>(null);
  //const [isDialogOpen, setIsDialogOpen] = useState<boolean>(false);
  //const [dialogType, setDialogType] = useState<"error" | "success" | "warning" | "info">("info"); // Default type

  //const client = context.aadHttpClientFactory.getClient("https://graph.microsoft.com");
  //const aadClient = context.aadHttpClientFactory.getClient("https://graph.microsoft.com");

  // Max Prod Team
  //const teamName = "ExpensesChat";
  //const channelName = "General";
  //const teamID = "68d9eb2c-06f7-40ed-bd99-a5a35fab0275";
  //const channelID = "19:WELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1@thread.tacv2";
  //const tagID = "NTA3NGI4Y2MtMTYwOC00YjQxLWFhZmQtMjY2MmRkNWY5YmZiIy…3LTQwZWQtYmQ5OS1hNWEzNWZhYjAyNzUjI3RndlFsV3dmTg==";
  //https://teams.microsoft.com/l/channel/19%3AWELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1%40thread.tacv2/General?groupId=68d9eb2c-06f7-40ed-bd99-a5a35fab0275&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb
  //https://teams.microsoft.com/l/team/19%3AWELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1%40thread.tacv2/conversations?groupId=68d9eb2c-06f7-40ed-bd99-a5a35fab0275&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb
  //https://teams.microsoft.com/l/team/19%3AWELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1%40thread.tacv2/conversations?groupId=68d9eb2c-06f7-40ed-bd99-a5a35fab0275&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb

  
  //const teamName = "Teams Testing";
  const teamID = "a3cce0fc-52f7-4928-8f2b-14102e5ad6ca";
  const channelID = "19:wREFwWCHiIj-qfeAUqedf6wIatZTFqg0CgOwMN6CQxc1@thread.tacv2";
  //const tagID = "NTA3NGI4Y2MtMTYwOC00YjQxLWFhZmQtMjY2MmRkNWY5YmZiIy…3LTQ5MjgtOGYyYi0xNDEwMmU1YWQ2Y2EjI3RMRnI2cXpQdw==";
  //https://teams.microsoft.com/l/team/19%3AwREFwWCHiIj-qfeAUqedf6wIatZTFqg0CgOwMN6CQxc1%40thread.tacv2/conversations?groupId=a3cce0fc-52f7-4928-8f2b-14102e5ad6ca&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb
  //https://teams.microsoft.com/l/channel/19%3AwREFwWCHiIj-qfeAUqedf6wIatZTFqg0CgOwMN6CQxc1%40thread.tacv2/General?groupId=a3cce0fc-52f7-4928-8f2b-14102e5ad6ca&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb

  // Max Dev Team
  //const teamName = "TestChat";
  //const teamID = "696dfe67-e76f-4bf8-8ab6-8abfcb16552e";
  //const channelName = "General";
  //https://teams.microsoft.com/l/channel/19%3AWELxtb3PBurFUqD2tVetv08tqw2FzQqvWFIqgi3XO5E1%40thread.tacv2/General?groupId=68d9eb2c-06f7-40ed-bd99-a5a35fab0275&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb
  //https://teams.microsoft.com/l/channel/19%3Aec62a56976504a9da063458459e73b34%40thread.tacv2/General?groupId=f5de9aad-7f98-498d-a5a0-b1a59254265c&tenantId=5074b8cc-1608-4b41-aafd-2662dd5f9bfb

  const userEmail = props.context.pageContext.user.email;

 /*
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
*/

  // Function to check if the user is a member of the team
  const isUserTeamMember = async (): Promise<boolean> => {
    try {
      const client = await context.msGraphClientFactory.getClient('3');
      const response = await client.api(`/groups/${teamID}/members`).get();
      const members = response.value || [];
      const user = await client.api('/me').get();
      const userId = user.id;

      const userIsMember = members.some((member: any) => member.id === userId);
      console.log("Is user a member of the team:", userIsMember);
      return userIsMember;
    } catch (error) {
      console.error("Error checking team membership:", error);
      return false;
    }
  };

  const getTeamTags = async (): Promise<void> => {  
    alert("Fetching tags for team ID: "+teamID);
    try {
      const client = await context.msGraphClientFactory.getClient('3');
      const response = await client.api(`/teams/${teamID}/tags`).get();
      if (response && response.value) {
        setTags(response.value);
        console.log('Tags:', response.value);
      }
    } catch (error) {
      console.error('Error getting MSGraphClient:', error);
    }

    /*
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
            console.log('Tags:', response.value);
          });
      });
    */
  
   return;
  };

  
  const sendMessageToTeams = async () : Promise<void> => {
    try {

      const aadClient = await context.aadHttpClientFactory.getClient("https://graph.microsoft.com");
      //const client = await context.msGraphClientFactory.getClient('3');
      const message = "Hello from the web part!";

      //console.log(client);
      // Fetch Team ID
      //const teamsResponse: HttpClientResponse = await aadClient.get(
      //  `https://graph.microsoft.com/v1.0/me/joinedTeams`,
      //  AadHttpClient.configurations.v1
      //);
      //if (!teamsResponse.ok) throw new Error("Failed to fetch teams");
  
      //const teamsData = await teamsResponse.json();
      //const team = teamsData.value.find((t: any) => t.displayName === teamName);
      //if (!team) throw new Error(`Team "${teamName}" not found`);
  
      // Fetch Team Tags (To get @expenses Tag ID)
      //const tagsResponse = await client.api(`/teams/${teamID}/tags`).get();
      
      //if (!tagsResponse.ok) throw new Error("Failed to fetch tags");

      //const tagsData = await tagsResponse.json();
      //const expensesTag = tagsData.value.find((tag: any) => tag.displayName === "expenses");
      console.log("Expenses Tag:", tags[0].id);

      //if (!expensesTag) throw new Error(`Tag "@expenses" not found in team "${teamName}"`);

      // Fetch Channel ID
      //const channelsResponse: HttpClientResponse = await client.get(
      //  `https://graph.microsoft.com/v1.0/teams/${teamID}/channels`,
      //  AadHttpClient.configurations.v1
      //);
      //if (!channelsResponse.ok) throw new Error("Failed to fetch channels");
  
      //const channelsData = await channelsResponse.json();
      //const channel = channelsData.value.find((c: any) => c.displayName === channelName);
      //if (!channel) throw new Error(`Channel "${channelName}" not found`);

      //console.log("sendmsg Team:", teamName);
      //console.log("sendmsg Channel:", channel);
  
      // 🔥 POST request to send message with @expenses mention

      const mentionId = 1; // You can keep this as 0 or another unique identifier, but it must match the ID in the <at> tag.
      const tagHTML = "<at id='1'>Expenses</at> ";

      console.log("Request Payload:", {
        body: {
          contentType: "html",
          content: tagHTML + message,
        },
        mentions: [
          {
            id: mentionId,
            mentionText: "Expenses",
            mentioned: {
              tag: {
                id: tags[0].id,
                displayName: "Expenses",
              },
            },
          },
        ],
      });

      const response = await aadClient.post(
        `https://graph.microsoft.com/v1.0/teams/${teamID}/channels/${channelID}/messages`,
        AadHttpClient.configurations.v1,
        {
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            body: {
              contentType: "html",
              content: tagHTML+message,
            },
            mentions: [
              {
                id: mentionId,
                mentionText: "expenses",
                mentioned: {
                  tag: {
                    id: tags[0].id,
                    displayName: "expenses",
                  },
                },
              },
            ],
          }),
        }
      );

      console.log("Mention ID:", mentionId);
  
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to send message: ${errorText}`);
      }

    } catch (error: any) {
      console.error("Error sending message:", error.message);
    }
    return;
  }
  
  const addMember = async (): Promise<void> => {
    const client = await context.msGraphClientFactory.getClient('3');
    const user = await client.api('/me').get();
    const userId = user.id;
  
    try {
      const apiUrl = `/teams/${teamID}/channels/${channelID}/members`; // Adding to a Teams channel chat
      const requestBody = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": ["member"], // Specify roles if needed
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userId}`
      };
  
      // Use the Graph Client V3 to make the API call
      await client.api(apiUrl).post(requestBody);
  
      alert('User added to the channel successfully!');
    } catch (error) {
      console.error('Error adding user to the channel:', error);
      alert(`Failed to add user to the channel. Error: ${error.message}`);
    }
  };
  
/*  
  const addMember = async (): Promise<void> => {
    const client = await context.msGraphClientFactory.getClient('3');
    const user = await client.api('/me').get();
    const userId = user.id;

    try {
      const apiUrl = `/groups/${teamID}/members/$ref`;
      const requestBody = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": ["member"], // Specify roles if needed, e.g., ["owner"]
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users/${userId}`,
        "visibleHistoryStartDateTime": new Date().toISOString() // Set desired start date for visible history
      };

      // Custom headers including x-ms-throttle-priority
      const customHeaders = {
        "x-ms-throttle-priority": "High"
      };

      // Use the Graph Client V3 to make the API call with custom headers
      await client
        .api(apiUrl)
        .headers(customHeaders) // Add the custom headers here
        .post(requestBody);

      alert('User added to the team successfully with high priority');
      
    } catch (error) {
      console.error('Error adding user to the team:', error);
      alert(`Failed to add user to the team. Error: ${error.message}`);
    }
  };  
*/

  // Function to ensure user is a member before fetching tags
  const ensureMembershipAndFetchTags = async (): Promise<void> => {
    const userIsMember = await isUserTeamMember();
    if (userIsMember) {
      console.log("User is a member. Fetching tags...");
      await getTeamTags();
    } else {
      console.log("User is not a member. Adding to the team...");
      //await addMember();
      // Optionally fetch tags after adding the user
      //setTimeout(async () => {
      //  await getTeamTags();
      //}, 3000); // Delay to ensure the user is added before fetching tags
    }
  };  

  useEffect(() => {

    const fetchGroupMembers = async ():Promise<void> => {
      try {
        const client: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
        const response = await client.api(`/groups/${teamID}/members`).get();

        if (response && response.value) {
          setGroupMembers(response.value);
          //getTeamTags();
        } else {
          console.warn('No group members found.');
        }
      } catch (error) {
        console.error('Error fetching group members:', error);
      }
    };

    fetchGroupMembers();
    ensureMembershipAndFetchTags(); 
    //.then(async() => {
      //setTimeout(async() => {
    //    await getTeamTags(); // Fetch tags after adding the user
        //await sendMessageToTeams("Hello from the web part!"); // Send a message to the Teams channel          
      //}, 3000); // Delay to ensure user is added before fetching tags  
    //});

  }, [context]);

  useEffect(() => {
    console.log("Group members updated:", groupMembers);
  }, [groupMembers]);
  
  useEffect(() => {
    console.log("Tags updated:", tags);
  }, [tags]);

  return (
    <section className={`${styles.testGraphClient} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
        <div>User Email: {escape(userEmail)}</div>
      </div>
      <div>
        <button onClick={addMember}>Join Chat</button>     
        <button onClick={getTeamTags}>Get Tags</button>
        <button onClick={sendMessageToTeams}>Post Msg</button>
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


/*

**** add member function ****

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
    
*******

        <div>
          {isDialogOpen && <Dialog type={dialogType} message={dialogMessage} onClose={() => setIsDialogOpen(false)} />}
        </div>


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

*/