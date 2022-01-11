import {
  AuthenticationProvider,
  AuthenticationProviderOptions,
} from "@microsoft/microsoft-graph-client";

import axios from "axios";

type MyAuthProviderOptions = {
  tenantID: string;
  clientSecret: string;
  clientID: string;
};

type TokenResponse = {
  token_type: string;
  expires_in: number;
  ext_expires_in: number;
  access_token: string;
};

class MyAuthProvider implements AuthenticationProvider {
  config: MyAuthProviderOptions;

  constructor(options: MyAuthProviderOptions) {
    this.config = options;
  }
  async getAccessToken(options: AuthenticationProviderOptions) {
    const params = new URLSearchParams();
    params.append("scope", "https://graph.microsoft.com/.default");
    params.append("client_id", this.config.clientID);
    params.append("client_secret", this.config.clientSecret);
    params.append("grant_type", "client_credentials");

    const cfg = {
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
    };

    const res = await axios.post<TokenResponse>(
      `https://login.microsoftonline.com/${this.config.tenantID}/oauth2/v2.0/token`,
      params,
      cfg
    );

    return res.data.access_token;
  }
}

export default MyAuthProvider;
