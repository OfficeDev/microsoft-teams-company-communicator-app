import React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { getBaseUrl } from './configVariables';

export interface IConfigState {
    url: string;
}

class Configuration extends React.Component<{}, IConfigState> {
    constructor(props: {}) {
        super(props);
        this.state = {
            url: getBaseUrl() + "/messages"
        }
    }

    public componentDidMount() {
        microsoftTeams.initialize();

        microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {
            microsoftTeams.settings.setSettings({
                entityId: "Company_Communicator_App",
                contentUrl: this.state.url,
                suggestedDisplayName: "Company Communicator Messages",
                websiteUrl: this.state.url,
            });
            saveEvent.notifySuccess();
        });

        microsoftTeams.settings.setValidityState(true);

    }

    public render() {
        return (
            <div>
                <h3>Company Communicator App Configuration Page</h3>
            </div>
        );
    }

}

export default Configuration;
