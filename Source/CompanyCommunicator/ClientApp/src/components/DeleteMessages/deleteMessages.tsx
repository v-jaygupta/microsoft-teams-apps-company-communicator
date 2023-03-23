// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { withTranslation, WithTranslation } from "react-i18next";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Button, Flex, RadioGroup, Datepicker } from '@fluentui/react-northstar';
import * as microsoftTeams from "@microsoft/teams-js";
import './deleteMessages.scss';
import { deleteHistoricalData } from '../../apis/messageListApi';
import { getBaseUrl } from '../../configVariables';
import { TFunction } from "i18next";

export interface IDeleteMessage {
    title?: string;
    height?: number;
    width?: number;
    url?: string;
    fallbackUrl?: string;
}

export interface formState {
    selectedRadioBtn: string,
    startDate: any,
    endDate: any,
    showDeleteTaskModule: boolean
}

export interface IDeleteMessagesProps extends WithTranslation {
    backNavigation?: any;
}

class DeleteMessages extends React.Component<IDeleteMessagesProps, formState> {
    readonly localize: TFunction;
    private isOpenTaskModuleAllowed: boolean;

    constructor(props: IDeleteMessagesProps) {
        super(props);
        initializeIcons();
        this.localize = this.props.t;
        this.isOpenTaskModuleAllowed = true;

        //Set default state values
        this.state = {
            selectedRadioBtn: "selectACustomDate",
            startDate: new Date(),
            endDate: new Date(),
            showDeleteTaskModule: false
        }
    }

    // Calculate particular time range based on selection
    private timeRange = [
        {
            name: "last30Days",
            value: [
                new Date(new Date().setDate(new Date().getDate() - 29)).toDateString(),
                new Date(new Date()).toDateString(),
            ]
        },
        {
            name: "last3Months",
            value: [
                new Date(new Date().getFullYear(), new Date().getMonth() - 3, 1).toDateString(),
                new Date(new Date().getFullYear(), new Date().getMonth(), 0).toDateString(),
            ]
        },
        {
            name: "last6Months",
            value: [
                new Date(new Date().getFullYear(), new Date().getMonth() - 6, 1).toDateString(),
                new Date(new Date().getFullYear(), new Date().getMonth(), 0).toDateString(),
            ]
        },
        {
            name: "last1Year",
            value: [
                new Date(new Date().getFullYear(), new Date().getMonth() - 12, 1).toDateString(),
                new Date(new Date().getFullYear(), new Date().getMonth(), 0).toDateString(),
            ]
        }
    ]

    // Handle radio button time range selection
    private onGroupSelected = (event: any, data: any) => {
        this.setState({
            selectedRadioBtn: data.value,
        });
    }

    // Back button navigation
    private onBack = (event: any) => {
        this.props.backNavigation(false);
    }

    // Send selected time range data on confirmation popup view
    private onApply = () => {
        let selectedTimeRange = null;
        let selectedTimeRangeName = '';
        if (this.state.selectedRadioBtn == "selectACustomDate") {
            selectedTimeRange = this.state.startDate.toDateString() + ',' + this.state.endDate.toDateString();
            selectedTimeRangeName = 'Custom Date';
        }
        else {
            selectedTimeRange = this.timeRange.filter(element => element.name == this.state.selectedRadioBtn)[0].value;
            selectedTimeRangeName = this.timeRange.filter(element => element.name == this.state.selectedRadioBtn)[0].name;
        }

        //Open confirmation popup to delete the selected time range data
        let url = getBaseUrl() + "/deletemessagestaskmodule/" + selectedTimeRange + "/" + selectedTimeRangeName + "?locale={locale}";
        this.onOpenTaskModule(null, url, this.localize("DeleteMessagesSectionTitle"));

    }

    // Open delete messages confirmation popup view
    public onOpenTaskModule = (event: any, url: string, title: string) => {

        if (this.isOpenTaskModuleAllowed) {
            this.isOpenTaskModuleAllowed = false;
            let taskInfo: IDeleteMessage = {
                url: url,
                title: title,
                height: 400,
                width: 490,
                fallbackUrl: url
            }

            let submitHandler = (err: any, result: any) => {
                this.isOpenTaskModuleAllowed = true;
            };

            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
        }
    }

    // Handle change of selected custom start date
    private onCutomStartDateHandler = (event: any, args: any) => {
        this.setState({
            startDate: args.itemProps.value.originalDate
        })
    }

    // Handle change of selected custom end date
    private onCutomEndDateHandler = (event: any, args: any) => {
        this.setState({
            endDate: args.itemProps.value.originalDate
        })
    }

    public render(): JSX.Element {
        return (
            <Flex className="deleteMessages">

                <Flex.Item size="size.full">
                    <Flex column className="formContentContainer">
                        <h4>{this.localize("ChooseRangeOfDeleteMessagesTitle")}</h4>

                        <RadioGroup
                            className="radioBtns"
                            checkedValue={this.state.selectedRadioBtn}
                            onCheckedValueChange={this.onGroupSelected}
                            vertical={true}
                            items={[
                                {
                                    name: "last30Days",
                                    key: "last30Days",
                                    value: "last30Days",
                                    label: this.localize("Last30Days"),
                                    children: (Component, { name, ...props }) => {
                                        return (
                                            <Flex key={name} column>
                                                <Component {...props} />
                                            </Flex>
                                        )
                                    },
                                },
                                {
                                    name: "last3Months",
                                    key: "last3Months",
                                    value: "last3Months",
                                    label: this.localize("Last3Months"),
                                    children: (Component, { name, ...props }) => {
                                        return (
                                            <Flex key={name} column>
                                                <Component {...props} />
                                            </Flex>
                                        )
                                    },
                                },
                                {
                                    name: "last6Months",
                                    key: "last6Months",
                                    value: "last6Months",
                                    label: this.localize("Last6Months"),
                                    children: (Component, { name, ...props }) => {
                                        return (
                                            <Flex key={name} column>
                                                <Component {...props} />
                                            </Flex>
                                        )
                                    },
                                },
                                {
                                    name: "last1Year",
                                    key: "last1Year",
                                    value: "last1Year",
                                    label: this.localize("Last1Year"),
                                    children: (Component, { name, ...props }) => {
                                        return (
                                            <Flex key={name} column>
                                                <Component {...props} />
                                            </Flex>
                                        )
                                    },
                                },
                                {
                                    name: "selectACustomDate",
                                    key: "selectACustomDate",
                                    value: "selectACustomDate",
                                    label: this.localize("SelectACustomDate"),
                                    children: (Component, { name, ...props }) => {
                                        return (
                                            <Flex key={name} column>
                                                <Component {...props} />

                                                <div className={this.state.selectedRadioBtn === "selectACustomDate" ? "displayCustomDatePicker" : "hide"}>
                                                    <div className="cutomDateRange">
                                                        <span>{this.localize("From")}</span>
                                                        <Datepicker
                                                            defaultSelectedDate={this.state.startDate}
                                                            onDateChange={this.onCutomStartDateHandler.bind(this)}
                                                        />
                                                    </div>
                                                    <hr />
                                                    <div className="cutomDateRange">
                                                        <span>{this.localize("To")}</span>
                                                        <Datepicker
                                                            defaultSelectedDate={this.state.endDate}
                                                            onDateChange={this.onCutomEndDateHandler.bind(this)}
                                                        />
                                                    </div>
                                                </div>
                                            </Flex>
                                        )
                                    },
                                }
                            ]}
                        >

                        </RadioGroup>
                    </Flex>
                </Flex.Item>
                <Flex className="footerContainer" vAlign="end" hAlign="end" gap="gap.small">
                    <Flex.Item push>
                        <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                    </Flex.Item>
                    <Button content={this.localize("Apply")} onClick={this.onApply} primary />
                </Flex>

            </Flex>
        );
    }
}

const newMessageWithTranslation = withTranslation()(DeleteMessages);
export default newMessageWithTranslation;