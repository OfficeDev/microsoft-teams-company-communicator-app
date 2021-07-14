// <copyright file="scheduledMessageOverflow.tsx" company="Microsoft Corporation">
// Copyright (c) Microsoft.
// Licensed under the MIT License.
// </copyright>

import React from 'react';
import { connect } from 'react-redux';
import { withTranslation, WithTranslation } from "react-i18next";
import { Menu, MoreIcon } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";

import { getBaseUrl } from '../../configVariables';
import { selectMessage, getMessagesList, getScheduledMessagesList, getDraftMessagesList } from '../../actions';
import { deleteDraftNotification, duplicateDraftNotification, sendPreview } from '../../apis/messageListApi';
import { TFunction } from "i18next";

export interface OverflowProps extends WithTranslation {
    message: any;
    styles?: object;
    title?: string;
    selectMessage?: any;
    getMessagesList?: any;
    getScheduledMessagesList?: any;
    getDraftMessagesList?: any;
}

export interface OverflowState {
    teamsTeamId?: string;
    teamsChannelId?: string;
    menuOpen: boolean;
}

export interface ITaskInfo {
    title?: string;
    height?: number;
    width?: number;
    url?: string;
    card?: string;
    fallbackUrl?: string;
    completionBotId?: string;
}

class Overflow extends React.Component<OverflowProps, OverflowState> {
    readonly localize: TFunction;
    constructor(props: OverflowProps) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            teamsChannelId: '',
            teamsTeamId: '',
            menuOpen: false,
        };
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.setState({
                teamsTeamId: context.teamId,
                teamsChannelId: context.channelId,
            });
        });
    }

    public render(): JSX.Element {
        const items = [
            {
                key: 'more',
                icon: <MoreIcon outline={true} />,
                menuOpen: this.state.menuOpen,
                active: this.state.menuOpen,
                indicator: false,
                menu: {
                    items: [
                        {
                            key: 'send',
                            content: this.localize("SendNow"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                let url = getBaseUrl() + "/sendconfirmation/" + this.props.message.id + "?locale={locale}";
                                this.onOpenTaskModule(null, url, this.localize("SendConfirmation"));
                            }
                        },
                        {
                            key: 'preview',
                            content: this.localize("PreviewInThisChannel"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                let payload = {
                                    draftNotificationId: this.props.message.id,
                                    teamsTeamId: this.state.teamsTeamId,
                                    teamsChannelId: this.state.teamsChannelId,
                                }
                                sendPreview(payload).then((response) => {
                                    return response.status;
                                }).catch((error) => {
                                    return error;
                                });
                            }
                        },
                        {
                            key: 'edit',
                            content: this.localize("Edit"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                let url = getBaseUrl() + "/newmessage/" + this.props.message.id + "?locale={locale}";
                                this.onOpenTaskModule(null, url, this.localize("EditMessage"));
                            }
                        },
                        {
                            key: 'duplicate',
                            content: this.localize("Duplicate"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                this.duplicateDraftMessage(this.props.message.id).then(() => {
                                    this.props.getDraftMessagesList();
                                });
                            }
                        },
                        {
                            key: 'delete',
                            content: this.localize("Delete"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                this.deleteDraftMessage(this.props.message.id).then(() => {
                                    this.props.getScheduledMessagesList(); 
                                });
                            }
                        },
                    ],
                },
                onMenuOpenChange: (e: any, { menuOpen }: any) => {
                    this.setState({
                        menuOpen: !this.state.menuOpen
                    });
                },
            },
        ];

        return <Menu className="menuContainer" iconOnly items={items} styles={this.props.styles} title={this.props.title} />;
    }

    private onOpenTaskModule = (event: any, url: string, title: string) => {
        let taskInfo: ITaskInfo = {
            url: url,
            title: title,
            height: 530,
            width: 1000,
            fallbackUrl: url,
        };

        let submitHandler = (err: any, result: any) => {
            this.props.getScheduledMessagesList().then(() => {
                this.props.getMessagesList();
            });
        };

        microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    }

    private duplicateDraftMessage = async (id: number) => {
        try {
            await duplicateDraftNotification(id);
        } catch (error) {
            return error;
        }
    }

    private deleteDraftMessage = async (id: number) => {
        try {
            await deleteDraftNotification(id);
        } catch (error) {
            return error;
        }
    }
}

const mapStateToProps = (state: any) => {
    return { messages: state.scheduledMessagesList, selectedMessage: state.selectedMessage };
}

const overflowWithTranslation = withTranslation()(Overflow);
export default connect(mapStateToProps, { selectMessage, getScheduledMessagesList, getDraftMessagesList, getMessagesList })(overflowWithTranslation);
