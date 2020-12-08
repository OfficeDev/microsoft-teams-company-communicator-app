import * as React from 'react';
import { connect } from 'react-redux';
import { withTranslation, WithTranslation } from "react-i18next";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { List } from '@stardust-ui/react';
import { Box, CircularProgress } from '@material-ui/core';
import * as microsoftTeams from "@microsoft/teams-js";

import './draftMessages.scss';
import { selectMessage, getDraftMessagesList, getMessagesList } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import Overflow from '../OverFlow/draftMessageOverflow';
import { TFunction } from "i18next";

export interface ITaskInfo {
  title?: string;
  height?: number;
  width?: number;
  url?: string;
  card?: string;
  fallbackUrl?: string;
  completionBotId?: string;
}

export interface IMessage {
  id: string;
  title: string;
  date: string;
  recipients: string;
  acknowledgements?: string;
  reactions?: string;
  responses?: string;
}

export interface IMessageProps extends WithTranslation {
  messages: IMessage[];
  selectedMessage: any;
  selectMessage?: any;
  getDraftMessagesList?: any;
  getMessagesList?: any;
}

export interface IMessageState {
  message: IMessage[];
  itemsAccount: number;
  CircularProgress: boolean;
  teamsTeamId?: string;
  teamsChannelId?: string;
}

class DraftMessages extends React.Component<IMessageProps, IMessageState> {
  readonly localize: TFunction;
  private interval: any;
  private isOpenTaskModuleAllowed: boolean;

  constructor(props: IMessageProps) {
    super(props);
    initializeIcons();
    this.localize = this.props.t;
    this.isOpenTaskModuleAllowed = true;
    this.state = {
      message: props.messages,
      itemsAccount: this.props.messages.length,
      CircularProgress: true,
      teamsTeamId: "",
      teamsChannelId: "",
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
    this.props.getDraftMessagesList();
    this.interval = setInterval(() => {
      this.props.getDraftMessagesList();
    }, 60000);
  }

  public componentWillReceiveProps(nextProps: any) {
    this.setState({
      message: nextProps.messages,
      CircularProgress: false
    })
  }

  public componentWillUnmount() {
    clearInterval(this.interval);
  }

  public render(): JSX.Element {
    let keyCount = 0;
    const processItem = (message: any) => {
      keyCount++;
      const out = {
        key: keyCount,
          content: (
              <div className="dFlex">
                  <Box>
                      <Box>
                        { message.title }
                      </Box>
                      <Box>
                        <Overflow message={ message } title="" />
                      </Box>
                </Box>
            </div>
        ),
        onClick: (): void => {
            let url = getBaseUrl() + "/newmessage/" + message.id + "?locale={locale}";
            this.onOpenTaskModule(null, url, this.localize("EditMessage"));
        },
      };
      return out;
    };

    const label = this.processLabels();
    const outList = this.state.message.map(processItem);
    const allDraftMessages = [...label, ...outList];

      if (this.state.CircularProgress) {
      return (
          <CircularProgress />
      );
    } else if (this.state.message.length === 0) {
        return (<div className="results">{this.localize("EmptyDraftMessages")}</div>);
    }
    else {
      return (
        <List selectable items={allDraftMessages} className="list" />
      );
    }
  }

  private processLabels = () => {
    const out = [{
      key: "labels",
        content: (
            <div>
                <Box>
                    { this.localize("TitleText") }
                </Box>
            </div>
      ),
    }];
    return out;
  }

  private onOpenTaskModule = (event: any, url: string, title: string) => {
    if (this.isOpenTaskModuleAllowed) {
      this.isOpenTaskModuleAllowed = false;
      let taskInfo: ITaskInfo = {
        url: url,
        title: title,
        height: 530,
        width: 1000,
        fallbackUrl: url,
      }

      let submitHandler = (err: any, result: any) => {
        this.props.getDraftMessagesList().then(() => {
          this.props.getMessagesList();
          this.isOpenTaskModuleAllowed = true;
        });
      };

      microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    }
  }
}

const mapStateToProps = (state: any) => {
  return { messages: state.draftMessagesList, selectedMessage: state.selectedMessage };
}

const draftMessagesWithTranslation = withTranslation()(DraftMessages);
export default connect(mapStateToProps, { selectMessage, getDraftMessagesList, getMessagesList })(draftMessagesWithTranslation);