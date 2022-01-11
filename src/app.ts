require("dotenv").config();

import "isomorphic-fetch";
import { ClientOptions, Client } from "@microsoft/microsoft-graph-client";
import { User, Group } from "@microsoft/microsoft-graph-types";
import { Team } from "@microsoft/microsoft-graph-types-beta";
import MyAuthProvider from "./myAuthProvider";

const clientSecret = process.env.CLIENT_SECRET as string;
const tenantID = process.env.TENANT_ID as string;
const clientID = process.env.CLIENT_ID as string;

console.log({ clientSecret, tenantID, clientID });

let clientOptions: ClientOptions = {
  authProvider: new MyAuthProvider({ tenantID, clientSecret, clientID }),
};

const client = Client.initWithMiddleware(clientOptions);

const getUsers = async (): Promise<User[]> =>
  (await client.api("/users").get()).value as User[];

const createGroup = async (users: User[]): Promise<Group> =>
  await client.api("/groups").post({
    displayName: "Test",
    mailNickname: "test",
    description: "This is a test",
    visibility: "Private",
    groupTypes: ["Unified"],
    mailEnabled: true,
    securityEnabled: false,
    "members@odata.bind": users
      .slice(0, -1)
      .map((user) => `https://graph.microsoft.com/v1.0/users/${user.id}`),
    "owners@odata.bind": [
      `https://graph.microsoft.com/v1.0/users/${users[users.length - 1].id}`,
    ],
  });

const createTeamFromGroup = async (groupId: string) =>
  await client.api(`/groups/${groupId}/team`).put({
    memberSettings: {
      allowCreatePrivateChannels: true,
      allowCreateUpdateChannels: true,
    },
    messagingSettings: {
      allowUserEditMessages: true,
      allowUserDeleteMessages: true,
    },
    funSettings: {
      allowGiphy: true,
      giphyContentRating: "strict",
    },
  });

const test = async () => {
  const users = (await client.api("/users").get()).value as User[];

  console.log(users);

  const groupId = (await createGroup(users)).id;

  console.log(groupId);

  // Group creation takes a while...
  //   const team = await createTeamFromGroup(groupId);
  //   console.log(team);
};

getUsers().then(console.log).catch(console.error);
