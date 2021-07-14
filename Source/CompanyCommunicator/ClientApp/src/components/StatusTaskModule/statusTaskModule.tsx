// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { withTranslation, WithTranslation } from "react-i18next";
import './statusTaskModule.scss';
import { getSentNotification, exportNotification } from '../../apis/messageListApi';
import { RouteComponentProps } from 'react-router-dom';
import * as AdaptiveCards from "adaptivecards";
import { TooltipHost } from 'office-ui-fabric-react';
import { Loader, List, Image, Button, DownloadIcon, AcceptIcon, Flex } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtns
} from '../AdaptiveCard/adaptiveCard';
import { ImageUtil } from '../../utility/imageutility';
import { formatDate, formatDuration, formatNumber } from '../../i18n';
import { TFunction } from "i18next";

export interface IListItem {
    header: string,
    media: JSX.Element,
}

export interface IMessage {
    id: string;
    title: string;
    acknowledgements?: string;
    reactions?: string;
    responses?: string;
    succeeded?: string;
    failed?: string;
    unknown?: string;
    sentDate?: string;
    imageLink?: string;
    summary?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
    teamNames?: string[];
    rosterNames?: string[];
    groupNames?: string[];
    allUsers?: boolean;
    sendingStartedDate?: string;
    sendingDuration?: string;
    errorMessage?: string;
    warningMessage?: string;
    canDownload?: boolean;
    sendingCompleted?: boolean;
    buttons: string;
    isImportant?: boolean;
}

export interface IStatusState {
    message: IMessage;
    loader: boolean;
    page: string;
    teamId?: string;
}

interface StatusTaskModuleProps extends RouteComponentProps, WithTranslation { }

class StatusTaskModule extends React.Component<StatusTaskModuleProps, IStatusState> {
    readonly localize: TFunction;
    private initMessage = {
        id: "",
        title: "",
        buttons: "[]",
    };

    private card: any;

    constructor(props: StatusTaskModuleProps) {
        super(props);

        this.localize = this.props.t;

        this.card = getInitAdaptiveCard(this.props.t);

        this.state = {
            message: this.initMessage,
            loader: true,
            page: "ViewStatus",
            teamId: '',
        };
    }

    public componentDidMount() {
        let params = this.props.match.params;
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.setState({
                teamId: context.teamId,
            });
        });

        if ('id' in params) {
            let id = params['id'];
            this.getItem(id).then(() => {
                this.setState({
                    loader: false
                }, () => {
                    setCardTitle(this.card, this.state.message.title);
                    setCardImageLink(this.card, this.state.message.imageLink);
                    setCardSummary(this.card, this.state.message.summary);
                    setCardAuthor(this.card, this.state.message.author);
                        
                    if (this.state.message.buttonTitle && this.state.message.buttonLink && !this.state.message.buttons) {
                        setCardBtns(this.card, [{
                            "type": "Action.OpenUrl",
                            "title": this.state.message.buttonTitle,
                            "url": this.state.message.buttonLink,
                        }]);
                        }
                        else {
                            setCardBtns(this.card, JSON.parse(this.state.message.buttons));
                    }

                    let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                    adaptiveCard.parse(this.card);
                    let renderedCard = adaptiveCard.render();
                    document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                    let link = this.state.message.buttonLink;
                    adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); }
                });
            });
        }
    }

    private getItem = async (id: number) => {
        try {
            const response = await getSentNotification(id);
            response.data.sendingDuration = formatDuration(response.data.sendingStartedDate, response.data.sentDate);
            response.data.sendingStartedDate = formatDate(response.data.sendingStartedDate);
            response.data.sentDate = formatDate(response.data.sentDate);
            response.data.succeeded = formatNumber(response.data.succeeded);
            response.data.failed = formatNumber(response.data.failed);
            response.data.unknown = response.data.unknown && formatNumber(response.data.unknown);
            this.setState({
                message: response.data
            });
        } catch (error) {
            return error;
        }
    }

    public render(): JSX.Element {
        if (this.state.loader) {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        } else {
            if (this.state.page === "ViewStatus") {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                            <Flex className="scrollableContent">
                                <Flex.Item size="size.half" className="formContentContainer">
                                    <Flex column>
                                        <div className="contentField">
                                            <h3>{this.localize("TitleText")}</h3>
                                            <span>{this.state.message.title}</span>
                                        </div>
                                        <div className="contentField">
                                            <h3>{this.localize("SendingStarted")}</h3>
                                            <span>{this.state.message.sendingStartedDate}</span>
                                        </div>
                                        <div className="contentField">
                                            <h3>{this.localize("Completed")}</h3>
                                            <span>{this.state.message.sentDate}</span>
                                        </div>
                                        <div className="contentField">
                                            <h3>{this.localize("Duration")}</h3>
                                            <span>{this.state.message.sendingDuration}</span>
                                        </div>
                                        <div className="contentField">
                                            <h3>{this.localize("Results")}</h3>
                                            <label>{this.localize("Success", { "SuccessCount": this.state.message.succeeded })}</label>
                                            <br />
                                            <label>{this.localize("Failure", { "FailureCount": this.state.message.failed })}</label>
                                            <br />
                                            {this.state.message.unknown &&
                                                <>
                                                    <label>{this.localize("Unknown", { "UnknownCount": this.state.message.unknown })}</label>
                                                </>
                                            }
                                        </div>
                                        <div className="contentField">
                                            <h3>{this.localize("Important")}</h3>
                                            <label>{this.renderImportant()}</label>
                                        </div>
                                        <div className="contentField">
                                            {this.renderAudienceSelection()}
                                        </div>
                                        <div className="contentField">
                                            {this.renderErrorMessage()}
                                        </div>
                                        <div className="contentField">
                                            {this.renderWarningMessage()}
                                        </div>
                                    </Flex>
                                </Flex.Item>
                                <Flex.Item size="size.half">
                                    <div className="adaptiveCardContainer">
                                    </div>
                                </Flex.Item>
                            </Flex>
                            <Flex className="footerContainer" vAlign="end" hAlign="end">
                                <div className={this.state.message.canDownload ? "" : "disabled"}>
                                    <Flex className="buttonContainer" gap="gap.small">
                                        <Flex.Item push>
                                            <Loader id="sendingLoader" className="hiddenLoader sendingLoader" size="smallest" label={this.localize("ExportLabel")} labelPosition="end" />
                                        </Flex.Item>
                                        <Flex.Item>
                                            <TooltipHost content={!this.state.message.sendingCompleted ? "" : (this.state.message.canDownload ? "" : this.localize("ExportButtonProgressText"))} calloutProps={{ gapSpace: 0 }}>
                                                <Button icon={<DownloadIcon size="medium" />} disabled={!this.state.message.canDownload || !this.state.message.sendingCompleted} content={this.localize("ExportButtonText")} id="exportBtn" onClick={this.onExport} primary />
                                            </TooltipHost>
                                        </Flex.Item>
                                    </Flex>
                                </div>
                            </Flex>
                        </Flex>
                    </div>
                );
            }
            else if (this.state.page === "SuccessPage") {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                            <div className="displayMessageField">
                                <br />
                                <br />
                                <div><span><AcceptIcon className="iconStyle" xSpacing="before" size="largest" outline /></span>
                                    <h1>{this.localize("ExportQueueTitle")}</h1></div>
                                <span>{this.localize("ExportQueueSuccessMessage1")}</span>
                                <br />
                                <br />
                                <span>{this.localize("ExportQueueSuccessMessage2")}</span>
                                <br />
                                <span>{this.localize("ExportQueueSuccessMessage3")}</span>
                            </div>
                            <Flex className="footerContainer" vAlign="end" hAlign="end" gap="gap.small">
                                <Flex className="buttonContainer">
                                    <Button content={this.localize("CloseText")} id="closeBtn" onClick={this.onClose} primary />
                                </Flex>
                            </Flex>
                        </Flex>
                    </div>
                );
            }
            else {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                            <div className="displayMessageField">
                                <br />
                                <br />
                                <div><span></span>
                                    <h1 className="light">{this.localize("ExportErrorTitle")}</h1></div>
                                <span>{this.localize("ExportErrorMessage")}</span>
                            </div>
                            <Flex className="footerContainer" vAlign="end" hAlign="end" gap="gap.small">
                                <Flex className="buttonContainer">
                                    <Button content={this.localize("CloseText")} id="closeBtn" onClick={this.onClose} primary />
                                </Flex>
                            </Flex>
                        </Flex>
                    </div>
                );
            }
        }
    }

    private onClose = () => {
        microsoftTeams.tasks.submitTask();
    }

    private onExport = async () => {
        let spanner = document.getElementsByClassName("sendingLoader");
        spanner[0].classList.remove("hiddenLoader");
        let payload = {
            id: this.state.message.id,
            teamId: this.state.teamId
        };
        await exportNotification(payload).then(() => {
            this.setState({ page: "SuccessPage" });
        }).catch(() => {
            this.setState({ page: "ErrorPage" });
        });
    }

    private getItemList = (items: string[]) => {
        let resultedTeams: IListItem[] = [];
        if (items) {
            resultedTeams = items.map((element) => {
                const resultedTeam: IListItem = {
                    header: element,
                    media: <Image src={ImageUtil.makeInitialImage(element)} avatar />
                }
                return resultedTeam;
            });
        }
        return resultedTeams;
    }

    private renderImportant = () => {
        if (this.state.message.isImportant) {
            return (
                <label>Yes</label>
            )
        } else {
            return (
                <label>No</label>
            )
        }
    }

    private renderAudienceSelection = () => {
        if (this.state.message.teamNames && this.state.message.teamNames.length > 0) {
            return (
                <div>
                    <h3>{this.localize("SentToGeneralChannel")}</h3>
                    <List items={this.getItemList(this.state.message.teamNames)} />
                </div>);
        } else if (this.state.message.rosterNames && this.state.message.rosterNames.length > 0) {
            return (
                <div>
                    <h3>{this.localize("SentToRosters")}</h3>
                    <List items={this.getItemList(this.state.message.rosterNames)} />
                </div>);
        } else if (this.state.message.groupNames && this.state.message.groupNames.length > 0) {
            return (
                <div>
                    <h3>{this.localize("SentToGroups1")}</h3>
                    <span>{this.localize("SentToGroups2")}</span>
                    <List items={this.getItemList(this.state.message.groupNames)} />
                </div>);
        } else if (this.state.message.allUsers) {
            return (
                <div>
                    <h3>{this.localize("SendToAllUsers")}</h3>
                </div>);
        } else {
            return (<div></div>);
        }
    }
    private renderErrorMessage = () => {
        if (this.state.message.errorMessage) {
            return (
                <div>
                    <h3>{this.localize("Errors")}</h3>
                    <span>{this.state.message.errorMessage}</span>
                </div>
            );
        } else {
            return (<div></div>);
        }
    }

    private renderWarningMessage = () => {
        if (this.state.message.warningMessage) {
            return (
                <div>
                    <h3>{this.localize("Warnings")}</h3>
                    <span>{this.state.message.warningMessage}</span>
                </div>
            );
        } else {
            return (<div></div>);
        }
    }
}

const StatusTaskModuleWithTranslation = withTranslation()(StatusTaskModule);
export default StatusTaskModuleWithTranslation;