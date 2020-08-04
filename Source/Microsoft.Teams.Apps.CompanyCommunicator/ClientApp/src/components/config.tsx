import React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { getBaseUrl } from '../configVariables';

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
                suggestedDisplayName: "Communicator Broadcaster Tab",
            });
            saveEvent.notifySuccess();
        });

        microsoftTeams.settings.setValidityState(true);

    }

    public render(): JSX.Element {
        return (
            <div className="configContainer">
                <h3>Please click Save to get started.</h3>
            </div>
        );
    }

}

export default Configuration;
