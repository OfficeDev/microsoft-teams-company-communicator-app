import * as React from 'react';
import { RouteComponentProps } from 'react-router-dom';
import { withTranslation, WithTranslation } from "react-i18next";
import { Input, TextArea, Radiobutton, RadiobuttonGroup } from 'msteams-ui-components-react';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import * as AdaptiveCards from "adaptivecards";
import { Button, Loader, Dropdown, Text } from '@stardust-ui/react';
import * as microsoftTeams from "@microsoft/teams-js";

import './newMessage.scss';
import './teamTheme.scss';
import { getDraftNotification, getTeams, createDraftNotification, updateDraftNotification, searchGroups, getGroups, verifyGroupAccess  } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtn
} from '../AdaptiveCard/adaptiveCard';
import { getBaseUrl } from '../../configVariables';
import { ImageUtil } from '../../utility/imageutility';
import { TFunction } from "i18next";

type dropdownItem = {
    key: string,
    header: string,
    content: string,
    image: string,
    team: {
        id: string,
    },
}

export interface IDraftMessage {
    id?: string,
    title: string,
    imageLink?: string,
    summary?: string,
    author: string,
    buttonTitle?: string,
    buttonLink?: string,
    teams: any[],
    rosters: any[],
    groups: any[],
    allUsers: boolean
}

export interface formState {
    title: string,
    summary?: string,
    btnLink?: string,
    imageLink?: string,
    btnTitle?: string,
    author: string,
    card?: any,
    page: string,
    teamsOptionSelected: boolean,
    rostersOptionSelected: boolean,
    allUsersOptionSelected: boolean,
    groupsOptionSelected: boolean,
    teams?: any[],
    groups?: any[],
    exists?: boolean,
    messageId: string,
    loader: boolean,
    groupAccess: boolean,
    loading: boolean,
    noResultMessage: string,
    unstablePinned?: boolean,
    selectedTeamsNum: number,
    selectedRostersNum: number,
    selectedGroupsNum: number,
    selectedRadioBtn: string,
    selectedTeams: dropdownItem[],
    selectedRosters: dropdownItem[],
    selectedGroups: dropdownItem[],
    errorImageUrlMessage: string,
    errorButtonUrlMessage: string,
    fileTitle?: string,
    file?: any,
}

export interface INewMessageProps extends RouteComponentProps, WithTranslation {
    getDraftMessagesList?: any;
}

class NewMessage extends React.Component<INewMessageProps, formState> {
    readonly localize: TFunction;
    private card: any;

    constructor(props: INewMessageProps) {
        super(props);
        initializeIcons();
        this.localize = this.props.t;
        this.card = getInitAdaptiveCard(this.localize);
        this.setDefaultCard(this.card);

        this.state = {
            title: "",
            summary: "",
            author: "",
            btnLink: "",
            imageLink: "",
            btnTitle: "",
            card: this.card,
            page: "CardCreation",
            teamsOptionSelected: true,
            rostersOptionSelected: false,
            allUsersOptionSelected: false,
            groupsOptionSelected: false,
            messageId: "",
            loader: true,
            groupAccess: false,
            loading: false,
            noResultMessage: "",
            unstablePinned: true,
            selectedTeamsNum: 0,
            selectedRostersNum: 0,
            selectedGroupsNum: 0,
            selectedRadioBtn: "teams",
            selectedTeams: [],
            selectedRosters: [],
            selectedGroups: [],
            errorImageUrlMessage: "",
            errorButtonUrlMessage: "",
            fileTitle: "",
            file: null,
        }
    }

    public async componentDidMount() {
        microsoftTeams.initialize();
        //- Handle the Esc key
        document.addEventListener("keydown", this.escFunction, false);
        let params = this.props.match.params;
        this.setGroupAccess();
        this.getTeamList().then(() => {
            if ('id' in params) {
                let id = params['id'];
                this.getItem(id).then(() => {
                    const selectedTeams = this.makeDropdownItemList(this.state.selectedTeams, this.state.teams);
                    const selectedRosters = this.makeDropdownItemList(this.state.selectedRosters, this.state.teams);
                    this.setState({
                        exists: true,
                        messageId: id,
                        selectedTeams: selectedTeams,
                        selectedRosters: selectedRosters,
                    })
                });
                this.getGroupData(id).then(() => {
                    const selectedGroups = this.makeDropdownItems(this.state.groups);
                    this.setState({
                        selectedGroups: selectedGroups
                    })
                });
            } else {
                this.setState({
                    exists: false,
                    loader: false
                }, () => {
                    let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                    adaptiveCard.parse(this.state.card);
                    let renderedCard = adaptiveCard.render();
                    document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                    if (this.state.btnLink) {
                        let link = this.state.btnLink;
                        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); };
                    }
                })
            }
        });
    }

    private makeDropdownItems = (items: any[] | undefined) => {
        const resultedTeams: dropdownItem[] = [];
        if (items) {
            items.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: {
                        id: element.id
                    },

                });
            });
        }
        return resultedTeams;
    }

    private makeDropdownItemList = (items: any[], fromItems: any[] | undefined) => {
        const dropdownItemList: dropdownItem[] = [];
        items.forEach(element =>
            dropdownItemList.push(
                typeof element !== "string" ? element : {
                    key: fromItems!.find(x => x.id === element).id,
                    header: fromItems!.find(x => x.id === element).name,
                    image: ImageUtil.makeInitialImage(fromItems!.find(x => x.id === element).name),
                    team: {
                        id: element
                    }
                })
        );
        return dropdownItemList;
    }

    public setDefaultCard = (card: any) => {
        const titleAsString = this.localize("TitleText");
        const summaryAsString = this.localize("Summary");
        const authorAsString = this.localize("Author1");
        const buttonTitleAsString = this.localize("ButtonTitle");

        setCardTitle(card, titleAsString);
        let imgUrl = getBaseUrl() + "/image/imagePlaceholder.png";
        setCardImageLink(card, imgUrl);
        setCardSummary(card, summaryAsString);
        setCardAuthor(card, authorAsString);
        setCardBtn(card, buttonTitleAsString, "https://adaptivecards.io");
    }

    private getTeamList = async () => {
        try {
            const response = await getTeams();
            this.setState({
                teams: response.data
            });
        } catch (error) {
            return error;
        }
    }

    private getGroupItems() {
        if (this.state.groups) {
            return this.makeDropdownItems(this.state.groups);
        }
        const dropdownItems: dropdownItem[] = [];
        return dropdownItems;
    }

    private setGroupAccess = async () => {
        await verifyGroupAccess().then(() => {
            this.setState({
                groupAccess: true
            });
        }).catch((error) => {
            const errorStatus = error.response.status;
            if (errorStatus === 403) {
                this.setState({
                    groupAccess: false
                });
            }
            else {
                throw error;
            }
        });
    }

    private getGroupData = async (id: number) => {
        try {
            const response = await getGroups(id);
            this.setState({
                groups: response.data
            });
        }
        catch (error) {
            return error;
        }
    }

    private getItem = async (id: number) => {
        try {
            const response = await getDraftNotification(id);
            const draftMessageDetail = response.data;
            let selectedRadioButton = "teams";
            if (draftMessageDetail.rosters.length > 0) {
                selectedRadioButton = "rosters";
            }
            else if (draftMessageDetail.groups.length > 0) {
                selectedRadioButton = "groups";
            }
            else if (draftMessageDetail.allUsers) {
                selectedRadioButton = "allUsers";
            }
            this.setState({
                teamsOptionSelected: draftMessageDetail.teams.length > 0,
                selectedTeamsNum: draftMessageDetail.teams.length,
                rostersOptionSelected: draftMessageDetail.rosters.length > 0,
                selectedRostersNum: draftMessageDetail.rosters.length,
                groupsOptionSelected: draftMessageDetail.groups.length > 0,
                selectedGroupsNum: draftMessageDetail.groups.length,
                selectedRadioBtn: selectedRadioButton,
                selectedTeams: draftMessageDetail.teams,
                selectedRosters: draftMessageDetail.rosters,
                selectedGroups: draftMessageDetail.groups
            });

            setCardTitle(this.card, draftMessageDetail.title);
            setCardImageLink(this.card, draftMessageDetail.imageLink);
            setCardSummary(this.card, draftMessageDetail.summary);
            setCardAuthor(this.card, draftMessageDetail.author);
            setCardBtn(this.card, draftMessageDetail.buttonTitle, draftMessageDetail.buttonLink);

            this.setState({
                title: draftMessageDetail.title,
                summary: draftMessageDetail.summary,
                btnLink: draftMessageDetail.buttonLink,
                imageLink: draftMessageDetail.imageLink,
                btnTitle: draftMessageDetail.buttonTitle,
                author: draftMessageDetail.author,
                allUsersOptionSelected: draftMessageDetail.allUsers,
                loader: false
            }, () => {
                this.updateCard();
            });
        } catch (error) {
            return error;
        }
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    public render(): JSX.Element {
        if (this.state.loader) {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        } else {
            if (this.state.page === "CardCreation") {
                return (
                    <div className="taskModule">
                        <div className="formContainer">
                            <div className="formContentContainer" >
                                <Input
                                    className="inputField"
                                    value={this.state.title}
                                    label={this.localize("TitleText")}
                                    placeholder={this.localize("PlaceHolderTitle")}
                                    onChange={this.onTitleChanged}
                                    autoComplete="off"
                                    required
                                />

                                <Input
                                    className="inputField"
                                    value={this.state.imageLink}
                                    label={this.localize("ImageURL")}
                                    placeholder={this.localize("ImageURL")}
                                    onChange={this.onImageLinkChanged}
                                    errorLabel={this.state.errorImageUrlMessage}
                                    autoComplete="off"
                                />

                                <TextArea
                                    className="inputField textArea"
                                    autoFocus
                                    placeholder={this.localize("Summary")}
                                    label={this.localize("Summary")}
                                    value={this.state.summary}
                                    onChange={this.onSummaryChanged}
                                />

                                <Input
                                    className="inputField"
                                    value={this.state.author}
                                    label={this.localize("Author")}
                                    placeholder={this.localize("Author")}
                                    onChange={this.onAuthorChanged}
                                    autoComplete="off"
                                />

                                <Input
                                    className="inputField"
                                    type="file"
                                    label={this.localize("File Attachment")}
                                    placeholder={this.localize("File Attachment")}
                                    onChange={this.onFileChanged}
                                    autoComplete="off"
                                />

                                <Input
                                    className="inputField"
                                    value={this.state.btnTitle}
                                    label={this.localize("ButtonTitle")}
                                    placeholder={this.localize("ButtonTitle")}
                                    onChange={this.onBtnTitleChanged}
                                    autoComplete="off"
                                />

                                <Input
                                    className="inputField"
                                    value={this.state.btnLink}
                                    label={this.localize("ButtonURL")}
                                    placeholder={this.localize("ButtonURL")}
                                    onChange={this.onBtnLinkChanged}
                                    errorLabel={this.state.errorButtonUrlMessage}
                                    autoComplete="off"
                                />
                            </div>
                            <div className="adaptiveCardContainer">
                            </div>
                        </div>

                        <div className="footerContainer">
                            <div className="buttonContainer">
                                <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                            </div>
                        </div>
                    </div>
                );
            }
            else if (this.state.page === "AudienceSelection") {
                return (
                    <div className="taskModule">
                        <div className="formContainer">
                            <div className="formContentContainer" >
                                <h3>{this.localize("SendHeadingText")}</h3>
                                <RadiobuttonGroup
                                    className="radioBtns"
                                    value={this.state.selectedRadioBtn}
                                    onSelected={this.onGroupSelected}
                                >
                                    <Radiobutton name="grouped" value="teams" label={this.localize("SendToGeneralChannel")} />
                                    <Dropdown
                                        hidden={!this.state.teamsOptionSelected}
                                        placeholder={this.localize("SendToGeneralChannelPlaceHolder")}
                                        search
                                        multiple
                                        items={this.getItems()}
                                        value={this.state.selectedTeams}
                                        onSelectedChange={this.onTeamsChange}
                                        noResultsMessage={this.localize("NoMatchMessage")}
                                    />
                                    <Radiobutton name="grouped" value="rosters" label={this.localize("SendToRosters")} />
                                    <Dropdown
                                        hidden={!this.state.rostersOptionSelected}
                                        placeholder={this.localize("SendToRostersPlaceHolder")}
                                        search
                                        multiple
                                        items={this.getItems()}
                                        value={this.state.selectedRosters}
                                        onSelectedChange={this.onRostersChange}
                                        unstable_pinned={this.state.unstablePinned}
                                        noResultsMessage={this.localize("NoMatchMessage")}
                                    />
                                    <Radiobutton name="grouped" value="allUsers" label={this.localize("SendToAllUsers")} />
                                    <div className={this.state.selectedRadioBtn === "allUsers" ? "" : "hide"}>
                                        <div className="noteText">
                                            <Text error content={this.localize("SendToAllUsersNote")} />
                                        </div>
                                    </div>
                                    <Radiobutton name="grouped" value="groups" label={this.localize("SendToGroups")} />
                                    <div className={this.state.groupsOptionSelected && !this.state.groupAccess ? "" : "hide"}>
                                        <div className="noteText">
                                            <Text error content={this.localize("SendToGroupsPermissionNote")} />
                                        </div>
                                    </div>
                                    <Dropdown
                                        className="hideToggle"
                                        hidden={!this.state.groupsOptionSelected || !this.state.groupAccess}
                                        placeholder={this.localize("SendToGroupsPlaceHolder")}
                                        search={this.onGroupSearch}
                                        multiple
                                        loading={this.state.loading}
                                        loadingMessage={this.localize("LoadingText")}
                                        items={this.getGroupItems()}
                                        value={this.state.selectedGroups}
                                        onSearchQueryChange={this.onGroupSearchQueryChange}
                                        onSelectedChange={this.onGroupsChange}
                                        noResultsMessage={this.state.noResultMessage}
                                        unstable_pinned={this.state.unstablePinned}
                                    />
                                    <div className={this.state.groupsOptionSelected && this.state.groupAccess ? "" : "hide"}>
                                        <div className="noteText">
                                            <Text error content={this.localize("SendToGroupsNote")} />
                                        </div>
                                    </div>
                                </RadiobuttonGroup>
                            </div>
                            <div className="adaptiveCardContainer">
                            </div>
                        </div>

                        <div className="footerContainer">
                            <div className="buttonContainer">
                                <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                                <Button content={this.localize("SaveAsDraft")} disabled={this.isSaveBtnDisabled()} id="saveBtn" onClick={this.onSave} primary />
                            </div>
                        </div>
                    </div>
                );
            } else {
                return (<div>Error</div>);
            }
        }
    }

    private onGroupSelected = (value: any) => {
        this.setState({
            selectedRadioBtn: value,
            teamsOptionSelected: value === 'teams',
            rostersOptionSelected: value === 'rosters',
            groupsOptionSelected: value === 'groups',
            allUsersOptionSelected: value === 'allUsers',
            selectedTeams: value === 'teams' ? this.state.selectedTeams : [],
            selectedTeamsNum: value === 'teams' ? this.state.selectedTeamsNum : 0,
            selectedRosters: value === 'rosters' ? this.state.selectedRosters : [],
            selectedRostersNum: value === 'rosters' ? this.state.selectedRostersNum : 0,
            selectedGroups: value === 'groups' ? this.state.selectedGroups : [],
            selectedGroupsNum: value === 'groups' ? this.state.selectedGroupsNum : 0,
        });
    }

    private isSaveBtnDisabled = () => {
        const teamsSelectionIsValid = (this.state.teamsOptionSelected && (this.state.selectedTeamsNum !== 0)) || (!this.state.teamsOptionSelected);
        const rostersSelectionIsValid = (this.state.rostersOptionSelected && (this.state.selectedRostersNum !== 0)) || (!this.state.rostersOptionSelected);
        const groupsSelectionIsValid = (this.state.groupsOptionSelected && (this.state.selectedGroupsNum !== 0)) || (!this.state.groupsOptionSelected);
        const nothingSelected = (!this.state.teamsOptionSelected) && (!this.state.rostersOptionSelected) && (!this.state.groupsOptionSelected) && (!this.state.allUsersOptionSelected);
        return (!teamsSelectionIsValid || !rostersSelectionIsValid || !groupsSelectionIsValid || nothingSelected)
    }

    private isNextBtnDisabled = () => {
        const title = this.state.title;
        const btnTitle = this.state.btnTitle;
        const btnLink = this.state.btnLink;
        return !(title && ((btnTitle && btnLink) || (!btnTitle && !btnLink)) && (this.state.errorImageUrlMessage === "") && (this.state.errorButtonUrlMessage === ""));
    }

    private getItems = () => {
        const resultedTeams: dropdownItem[] = [];
        if (this.state.teams) {
            let remainingUserTeams = this.state.teams;
            if (this.state.selectedRadioBtn !== "allUsers") {
                if (this.state.selectedRadioBtn === "teams") {
                    this.state.teams.filter(x => this.state.selectedTeams.findIndex(y => y.team.id === x.id) < 0);
                }
                else if (this.state.selectedRadioBtn === "rosters") {
                    this.state.teams.filter(x => this.state.selectedRosters.findIndex(y => y.team.id === x.id) < 0);
                }
            }
            remainingUserTeams.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: {
                        id: element.id
                    }
                });
            });
        }
        return resultedTeams;
    }

    private static MAX_SELECTED_TEAMS_NUM: number = 20;

    private onTeamsChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        this.setState({
            selectedTeams: itemsData.value,
            selectedTeamsNum: itemsData.value.length,
            selectedRosters: [],
            selectedRostersNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0
        })
    }

    private onRostersChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        this.setState({
            selectedRosters: itemsData.value,
            selectedRostersNum: itemsData.value.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0
        })
    }

    private onGroupsChange = (event: any, itemsData: any) => {
        this.setState({
            selectedGroups: itemsData.value,
            selectedGroupsNum: itemsData.value.length,
            groups: [],
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedRosters: [],
            selectedRostersNum: 0
        })
    }

    private onGroupSearch = (itemList: any, searchQuery: string) => {
        const result = itemList.filter(
            (item: { header: string; content: string; }) => (item.header && item.header.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1) ||
                (item.content && item.content.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1),
        )
        return result;
    }

    private onGroupSearchQueryChange = async (event: any, itemsData: any) => {

        if (!itemsData.searchQuery) {
            this.setState({
                groups: [],
                noResultMessage: "",
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length <= 2) {
            this.setState({
                loading: false,
                noResultMessage: "No matches found.",
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length > 2) {
            // handle event trigger on item select.
            const result = itemsData.items && itemsData.items.find(
                (item: { header: string; }) => item.header.toLowerCase() === itemsData.searchQuery.toLowerCase()
            )
            if (result) {
                return;
            }

            this.setState({
                loading: true,
                noResultMessage: "",
            });

            try {
                const query = encodeURIComponent(itemsData.searchQuery);
                const response = await searchGroups(query);
                this.setState({
                    groups: response.data,
                    loading: false,
                    noResultMessage: "No matches found."
                });
            }
            catch (error) {
                return error;
            }
        }
    }

    private onSave = () => {
        const selectedTeams: string[] = [];
        const selctedRosters: string[] = [];
        const selectedGroups: string[] = [];
        this.state.selectedTeams.forEach(x => selectedTeams.push(x.team.id));
        this.state.selectedRosters.forEach(x => selctedRosters.push(x.team.id));
        this.state.selectedGroups.forEach(x => selectedGroups.push(x.team.id));

        const draftMessage: IDraftMessage = {
            id: this.state.messageId,
            title: this.state.title,
            imageLink: this.state.imageLink,
            summary: this.state.summary,
            author: this.state.author,
            buttonTitle: this.state.btnTitle,
            buttonLink: this.state.btnLink,
            teams: selectedTeams,
            rosters: selctedRosters,
            groups: selectedGroups,
            allUsers: this.state.allUsersOptionSelected
        };

        const formData = new FormData();
        const draftMessageString = JSON.stringify(draftMessage)
        formData.append("draftMessage", draftMessageString);

        if (this.state.file) {
            formData.append("file", this.state.file, this.state.fileTitle);
        }

        if (this.state.exists) {
            this.editDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        } else {
            this.postDraftMessage(formData).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        }
    }

    private editDraftMessage = async (draftMessage: IDraftMessage) => {
        try {
            await updateDraftNotification(draftMessage);
        } catch (error) {
            return error;
        }
    }

    private postDraftMessage = async (draftMessage: any) => {
        try {
            await createDraftNotification(draftMessage);
        } catch (error) {
            throw error;
        }
    }

    public escFunction(event: any) {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }

    private onNext = (event: any) => {
        this.setState({
            page: "AudienceSelection"
        }, () => {
            this.updateCard();
        });
    }

    private onBack = (event: any) => {
        this.setState({
            page: "CardCreation"
        }, () => {
            this.updateCard();
        });
    }

    private onTitleChanged = (event: any) => {
        let showDefaultCard = (!event.target.value && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, event.target.value);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        this.setState({
            title: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onImageLinkChanged = (event: any) => {
        let url = event.target.value.toLowerCase();
        if (!((url === "") || (url.startsWith("https://") || (url.startsWith("data:image/png;base64,")) || (url.startsWith("data:image/jpeg;base64,")) || (url.startsWith("data:image/gif;base64,"))))) {
            this.setState({
                errorImageUrlMessage: "URL must start with https://"
            });
        } else {
            this.setState({
                errorImageUrlMessage: ""
            });
        }

        let showDefaultCard = (!this.state.title && !event.target.value && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, event.target.value);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        this.setState({
            imageLink: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onSummaryChanged = (event: any) => {
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !event.target.value && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, event.target.value);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        this.setState({
            summary: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onAuthorChanged = (event: any) => {
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !event.target.value && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, event.target.value);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        this.setState({
            author: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onBtnTitleChanged = (event: any) => {
        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnLink);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        if (event.target.value && this.state.btnLink) {
            setCardBtn(this.card, event.target.value, this.state.btnLink);
            this.setState({
                btnTitle: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            this.setState({
                btnTitle: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLinkChanged = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage: "URL must start with https://"
            });
        } else {
            this.setState({
                errorButtonUrlMessage: ""
            });
        }

        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !event.target.value);
        setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        if (this.state.btnTitle && event.target.value) {
            setCardBtn(this.card, this.state.btnTitle, event.target.value);
            this.setState({
                btnLink: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            this.setState({
                btnLink: event.target.value
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private updateCard = () => {
        const adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(this.state.card);
        const renderedCard = adaptiveCard.render();
        const container = document.getElementsByClassName('adaptiveCardContainer')[0].firstChild;
        if (container != null) {
            container.replaceWith(renderedCard);
        } else {
            document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
        }
        const link = this.state.btnLink;
        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); }
    }

    private onFileChanged = (event: any) => {
        this.setState({
            fileTitle: event.target.files[0].name,
            file: event.target.files[0]
        });
    }
}

const newMessageWithTranslation = withTranslation()(NewMessage);
export default newMessageWithTranslation;
