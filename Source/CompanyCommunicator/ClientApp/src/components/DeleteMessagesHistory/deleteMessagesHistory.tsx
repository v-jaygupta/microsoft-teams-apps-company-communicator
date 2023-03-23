// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { withTranslation, WithTranslation } from "react-i18next";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Loader, List, Flex, Text } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";
import { getDeleteMessagesList } from '../../actions';
import { getDeleteMessagesData } from '../../apis/messageListApi';
import { TFunction } from "i18next";
import { Icon, TooltipHost } from 'office-ui-fabric-react';

export interface IMessageProps extends WithTranslation {
    getDeleteMessagesList?: any;
    deleteMessagesProps: any;
}

export interface IMessageState {
    deleteMessage: any;
    loader: boolean;
}

class DeleteMessagesHistory extends React.Component<IMessageProps, IMessageState> {
    readonly localize: TFunction;
    private interval: any;

    constructor(props: IMessageProps) {
        super(props);
        initializeIcons();
        this.localize = this.props.t;
        this.state = {
            deleteMessage: [],
            loader: true
        };
    }
    public deleteHistoricalMessages = async () => {
        try {
            await getDeleteMessagesData().then((messageResponse) => {
                this.state = {
                    deleteMessage: messageResponse,
                    loader: false
                };
            }).catch((ex) => {
                console.log(ex)
            });
        }
        catch (error) {
            return error;
        }
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        this.deleteHistoricalMessages();
        getDeleteMessagesList();
        this.render();
        this.interval = setInterval(() => {
            getDeleteMessagesList();
        }, 60000);
    }
    public componentWillUnmount() {
        clearInterval(this.interval);
    }
    // public componentWillReceiveProps(nextProps: any) {
    //     debugger;
    //     if (this.props !== nextProps) {
    //         this.setState({
    //             deleteMessage: this.props.deleteMessagesProps.data,
    //             loader: false
    //         });
    //     }
    // }

    public tooltipContent(message: any) {
        return (<div>{this.localize("From")} <br /> {message.startDate} <br /> <hr /> {this.localize("To")} <br /> {message.endDate}</div>);

    }

    public render(): JSX.Element {
        let keyCount = 0;
        const processItem = (message: any) => {
            keyCount++;
            const out = {
                key: keyCount,
                content: this.messageContent(message),
                styles: { margin: '0.2rem 0.2rem 0 0' },
            };
            return out;
        };
        const label = this.processLabels();
        const outList = this.props.deleteMessagesProps && this.props.deleteMessagesProps.data.map(processItem);
        const allMessages = [...label, ...outList];

        // if (this.state.loader) {
        //     return (
        //         <Loader />
        //     );
        // } else if (this.props.deleteMessagesProps.data.length === 0) {
        //     return (<div className="results">{this.localize("EmptySentMessages")}</div>);
        // }
        if (this.props.deleteMessagesProps.data.length === 0) {
            return (<div className="results">{this.localize("EmptyDeleteMessages")}</div>);
        }
        else {
            return (
                <div className='deleteMessagesHistory'>
                    <List selectable items={allMessages} className="list" />
                </div>
            );
        }
    }

    private processLabels = () => {
        const out = [{
            key: "labels",
            content: (
                <Flex vAlign="center" fill gap="gap.small">
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} grow={1} >
                        <Text
                            truncated
                            weight="bold"
                            content={this.localize("Date&Time")}
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '16%' }}>
                        <Text
                            truncated
                            weight="bold"
                            content={this.localize("Status")}
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} >
                        <Text
                            truncated
                            content={this.localize("DataRange")}
                            weight="bold"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }} >
                        <Text
                            truncated
                            content={this.localize("NumberOfRecordsDeleted")}
                            weight="bold"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} >
                        <Text
                            truncated
                            content={this.localize("DeletedBy")}
                            weight="bold"
                        >
                        </Text>
                    </Flex.Item>
                </Flex>
            ),
            styles: { margin: '0.2rem 0.2rem 0 0' },
        }];
        return out;
    }

    private messageContent = (message: any) => {
        return (
            <Flex className="listContainer" vAlign="center" fill gap="gap.small">
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }} grow={1}>
                    <Text
                        truncated
                        content={new Date(message.timestamp).toDateString()}
                    >
                    </Text>
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '16%' }}>
                    <Text
                        truncated
                        content={message.status} />
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }} shrink={false}>
                    <div>
                        <TooltipHost content={this.tooltipContent(message)} hostClassName='tooltipHostStyles' calloutProps={{ gapSpace: 0 }}
                            tooltipProps={{
                                calloutProps: {
                                    styles: {
                                        calloutMain: { background: '#000' },
                                        root: { border: '1px solid #333', width: '15rem' },
                                        beak: { background: '#000' },
                                        beakCurtain: { background: '#000' },
                                    },
                                },
                                styles: {
                                    root: { background: '#000' },
                                    content: { background: '#000', padding: '0px 5px', color: '#fff' }
                                }
                            }}>
                            <Text
                                truncated
                                content={message.selectedDateRange} />
                            <Icon aria-label="Info" iconName="Info" className='tooltipHostStylesInsideContent' />
                        </TooltipHost>
                    </div>
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }}>
                    <Text
                        truncated
                        content={message.recordsDeleted} />
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '20%' }}>
                    <Text
                        truncated
                        content={message.deletedBy} />
                </Flex.Item>
            </Flex>
        );
    }
}
const draftMessagesWithTranslation = withTranslation()(DeleteMessagesHistory);
export default draftMessagesWithTranslation;