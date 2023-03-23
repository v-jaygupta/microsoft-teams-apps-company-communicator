// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { withTranslation, WithTranslation } from "react-i18next";
import Messages from '../Messages/messages';
import DraftMessages from '../DraftMessages/draftMessages';
import DeleteMessagesHistory from '../DeleteMessagesHistory/deleteMessagesHistory';
import DeleteMessages from '../DeleteMessages/deleteMessages';
import './tabContainer.scss';
import * as microsoftTeams from "@microsoft/teams-js";
import { getBaseUrl } from '../../configVariables';
import { Accordion, Button, Flex, FlexItem, Tooltip } from '@fluentui/react-northstar';
import { getDraftMessagesList } from '../../actions';
import { connect } from 'react-redux';
import { TFunction } from "i18next";
import { NavLink } from 'react-router-dom';
import { getDeleteMessagesData } from '../../apis/messageListApi';

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
    backNavigation: any;
    deleteMessage: any;
    deleteMessagesProps: any;
}

export interface ITabContainerState {
    url: string,
    isDeleteHistoryClicked: boolean,
    currentLoggedInUser: any,
    hasUserDeletePermission: boolean,
    deleteMessage: any
}

class TabContainer extends React.Component<ITaskInfoProps, ITabContainerState> {

    readonly localize: TFunction;
    constructor(props: ITaskInfoProps) {
        super(props);
        this.localize = this.props.t;

        this.state = {
            url: getBaseUrl() + "/newmessage?locale={locale}",
            isDeleteHistoryClicked: false,
            currentLoggedInUser: {},
            hasUserDeletePermission: false,
            deleteMessage: ''
        }
        this.escFunction = this.escFunction.bind(this);
    }

    // Check if the user has delete messages permission
    loggedInUserAuthentication = (currentLoggedInUser: any) => {

        // Get user mail id using environment variable
        let authorizedUsers = process.env.REACT_APP_AUTHORIZED_USERS_EMAIL;
        const deteleMessagePermission = authorizedUsers && authorizedUsers.toLowerCase().includes(currentLoggedInUser);

        this.setState({
            hasUserDeletePermission: deteleMessagePermission ? deteleMessagePermission : this.state.hasUserDeletePermission
        })
    }

    public deleteHistoricalMessages = async () => {
        try {
            debugger;
            const respone = await getDeleteMessagesData();
            this.setState({
                deleteMessage: respone
            });
        }
        catch (error) {
            return error;
        }
    }

    public componentDidMount() {
        microsoftTeams.initialize();

        //Get Current Logged in user detail
        microsoftTeams.getContext((context) => {
            const currentLoggedInUser = context.userPrincipalName ? context.userPrincipalName.toLowerCase() : "";
            this.loggedInUserAuthentication(currentLoggedInUser);
        });

        this.deleteHistoricalMessages();
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

    // Get back button click details from deleteMessages page
    public handleCallback = (isBackButtonClicked: any) => {
        this.setState({ isDeleteHistoryClicked: isBackButtonClicked })
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
                title: this.localize('SentMessagesSectionTitle'),
                content: {
                    key: 'draft',
                    content: (
                        <Messages></Messages>
                    ),
                },
            }
        ]
        // Define a new panel to delete selected time range messages
        const deletePanels = [
            {
                title: '',
                content: {
                    key: 'deleteMessages',
                    content: (
                        <DeleteMessages backNavigation={this.handleCallback}></DeleteMessages>
                    ),
                },
            },
            {
                title: this.localize('DeleteMessagesHistorySectionTitle'),
                content: {
                    key: 'deleteMessagesHistory',
                    content: (
                        <DeleteMessagesHistory deleteMessagesProps={this.state.deleteMessage}></DeleteMessagesHistory>
                    ),
                },
            }
        ]
        const buttonId = 'callout-button';
        const customHeaderImagePath = process.env.REACT_APP_HEADERIMAGE;
        const customHeaderText = process.env.REACT_APP_HEADERTEXT == null ?
            this.localize("CompanyCommunicator")
            : this.localize(process.env.REACT_APP_HEADERTEXT);

        return (
            <>
                <div className='cc-header'>
                    <Flex gap="gap.small" space='between'>
                        <Flex gap="gap.small" vAlign="center">
                            <img
                                src={customHeaderImagePath == null ?
                                    require("../../assets/Images/mslogo.png").default
                                    : customHeaderImagePath}
                                alt="Ms Logo"
                                className="ms-logo"
                                title={customHeaderText}
                            />
                            <span className="header-text" title={customHeaderText}>{customHeaderText}</span>
                        </Flex>
                        <Flex gap="gap.large" vAlign="center">
                            <FlexItem>
                                <a
                                    href="https://aka.ms/M365CCIssues"
                                    target="_blank" rel="noreferrer"
                                >
                                    <Tooltip
                                        trigger={<img
                                            src={require("../../assets/Images/HelpIcon.svg").default}
                                            alt="Help"
                                            className="support-icon"
                                        />}
                                        content={this.localize("Support")}
                                        pointing={false}
                                    />
                                </a>
                            </FlexItem>
                            <FlexItem>
                                <a href="https://aka.ms/M365CCFeedback" target="_blank" rel="noreferrer">
                                    <Tooltip
                                        trigger={<img
                                            src={require("../../assets/Images/FeedbackIcon.svg").default}
                                            alt="Feedback"
                                            className='feedback-icon'
                                        />}
                                        content={this.localize("Feedback")}
                                        pointing={false}
                                    />
                                </a>
                            </FlexItem>
                        </Flex>
                    </Flex>
                </div>
                <Flex className="tabContainer" column fill gap="gap.small">

                    {!this.state.isDeleteHistoryClicked
                        ? <>
                            <Flex className="newPostBtn" hAlign="end" vAlign="end">
                                 {this.state.hasUserDeletePermission ?
                                    <Button content={this.localize("DeleteMessages")} onClick={this.onDeleteMessages} secondary />
                                    : ''
                                }
                                <Button content={this.localize("NewMessage")} onClick={this.onNewMessage} primary />
                            </Flex>
                            <Flex className="messageContainer">
                                <Flex.Item grow={1} >
                                    <Accordion defaultActiveIndex={[0, 1]} panels={panels} />
                                </Flex.Item>
                            </Flex>
                        </>
                        : <Flex className="messageContainer">
                            <div className="topNavigation">
                                <nav className="navbar is-primary" role="navigation" aria-label="main navigation">{this.localize("Dashboard")}
                                    <NavLink to='deletemessages'>{this.localize("DeleteMessagesSectionTitle")}</NavLink>
                                </nav>
                            </div>

                            <Flex.Item grow={1} >
                                <Accordion defaultActiveIndex={[0, 1]} panels={deletePanels} />
                            </Flex.Item>
                        </Flex>
                    }
                </Flex>
            </>
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
        };

        microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    }

    onDeleteMessages = () => {

        this.setState({
            isDeleteHistoryClicked: true
        });
    }
}

const mapStateToProps = (state: any) => {
    return { messages: state.draftMessagesList };
}

const tabContainerWithTranslation = withTranslation()(TabContainer);
export default connect(mapStateToProps, { getDraftMessagesList })(tabContainerWithTranslation);