// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { withTranslation, WithTranslation } from "react-i18next";
import Messages from '../Messages/messages';
import DraftMessages from '../DraftMessages/draftMessages';
import ScheduledMessages from '../ScheduledMessages/ScheduledMessages';
import './tabContainer.scss';
import * as microsoftTeams from "@microsoft/teams-js";
import { getBaseUrl } from '../../configVariables';
import { Accordion, Button, Flex } from '@fluentui/react-northstar';
import { getDraftMessagesList, getScheduledMessagesList } from '../../actions';
import { connect } from 'react-redux';
import { TFunction } from "i18next";


interface ITaskInfo {
    title?: string;
    height?: number;
    width?: number;
    url?: string;
    card?: string;
    fallbackUrl?: string;
    completionBotId?: string;
}

export interface ITaskInfoProps extends WithTranslation {
    getDraftMessagesList?: any;
    getScheduledMessagesList?: any;
}

export interface ITabContainerState {
    url: string;
}

class TabContainer extends React.Component<ITaskInfoProps, ITabContainerState> {
    readonly localize: TFunction;
    constructor(props: ITaskInfoProps) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            url: getBaseUrl() + "/newmessage?locale={locale}"
        }
        this.escFunction = this.escFunction.bind(this);
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        //- Handle the Esc key
        document.addEventListener("keydown", this.escFunction, false);
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    public escFunction(event: any) {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }

    public render(): JSX.Element {
        const panels = [
            {
                title: this.localize('DraftMessagesSectionTitle'),
                content: {
                    key: 'sent',
                    content: (
                        <DraftMessages></DraftMessages>
                    ),
                },
            },
            {
                title: this.localize('ScheduledMessagesSectionTitle'),
                content: {
                    key: 'scheduled',
                    content: (
                        <div className="messages">
                            <ScheduledMessages></ScheduledMessages>
                        </div>
                    ),
                },
            },
            {
                title: this.localize('SentMessagesSectionTitle'),
                content: {
                    key: 'draft',
                    content: (
                        <Messages></Messages>
                    ),
                },
            }
        ]
        return (
            <Flex className="tabContainer" column fill gap="gap.small">
                <Flex className="newPostBtn" hAlign="end" vAlign="end">
                    <Button content={this.localize("NewMessage")} onClick={this.onNewMessage} primary />
                </Flex>
                <Flex className="messageContainer">
                    <Flex.Item grow={1} >
                        <Accordion defaultActiveIndex={[0, 1, 2]} panels={panels} />
                    </Flex.Item>
                </Flex>
            </Flex>
        );
    }

    public onNewMessage = () => {
        let taskInfo: ITaskInfo = {
            url: this.state.url,
            title: this.localize("NewMessage"),
            height: 530,
            width: 1000,
            fallbackUrl: this.state.url,
        }

        let submitHandler = (err: any, result: any) => {
            this.props.getDraftMessagesList();
            this.props.getScheduledMessagesList();
            
        };

        microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    }
}

const mapStateToProps = (state: any) => {
    return { messages: state.draftMessagesList };
}

const tabContainerWithTranslation = withTranslation()(TabContainer);
export default connect(mapStateToProps, { getDraftMessagesList, getScheduledMessagesList })(tabContainerWithTranslation);