// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from "react";
import { Text, Button, Flex, Loader } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { withTranslation, WithTranslation } from "react-i18next";
import { TFunction } from "i18next";
import { deleteHistoricalData } from '../../apis/messageListApi';
import { RouteComponentProps } from 'react-router-dom';
import './deleteMessagesTaskModule.scss';

export interface IDeleteTaskModuleProps extends RouteComponentProps, WithTranslation {
    deleteHistoricalMessages?: any;
}

export interface formState {
    selectedTimeRange: any,
    historicalData?: any[],
    loader: boolean,
    startDate: any,
    endDate:any,
    selectedTimeRangeName: any,
    currentLoggedInUser: any
}

class DeleteMessagesTaskModule extends React.Component<IDeleteTaskModuleProps, formState> {
    readonly localize: TFunction;
    constructor(props: IDeleteTaskModuleProps) {

        super(props);
        initializeIcons();
        this.localize = this.props.t;

        this.state = {
            selectedTimeRange: null,
            historicalData: [],
            loader: true,
            startDate: null,
            endDate: null,
            selectedTimeRangeName: '',
            currentLoggedInUser: ''
        }
    }

    public async componentDidMount() {
        let params = this.props.match.params;
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.setState({
                currentLoggedInUser : context.userPrincipalName ? context.userPrincipalName.toLowerCase() : ""
            });
        });
        if ('selectedTimeRange' in params && 'selectedTimeRangeName' in params ) {
            //Get selectedTimeRange from url parameters
            let selectedTimeRange:any = params['selectedTimeRange'];
            let selectedTimeRangeName:any = params['selectedTimeRangeName'];
            this.setState({
                selectedTimeRange: selectedTimeRange,
                loader: false,
                startDate: selectedTimeRange.split(',')[0],
                endDate: selectedTimeRange.split(',')[1],
                selectedTimeRangeName: selectedTimeRangeName
            })
        }
    }

    // Prepare API call to delete selected time range data
    public deleteHistoricalMessages = async () => {
        try {
            let selectedTimeRangeNameValue = '';
            if(this.state.selectedTimeRangeName == "last30Days")
            {
                selectedTimeRangeNameValue = this.localize("Last30Days");                   
            }
            else if(this.state.selectedTimeRangeName == "last3Months")
            {
                selectedTimeRangeNameValue = this.localize("Last3Months");                   
            }
            else if(this.state.selectedTimeRangeName == "last6Months")
            {
                selectedTimeRangeNameValue = this.localize("Last6Months");                   
            }
            else if (this.state.selectedTimeRangeName == "last1Year")
            {
                selectedTimeRangeNameValue = this.localize("Last1Year");                   
            }
            else if (this.state.selectedTimeRangeName == "Custom Date")
            {
                selectedTimeRangeNameValue = "Custom Date";                   
            }
            
            let payload = {
                selectedDateRange : selectedTimeRangeNameValue,
                deletedBy: this.state.currentLoggedInUser,
                startDate: this.state.startDate,
                endDate: this.state.endDate,
            };
            await deleteHistoricalData(payload).then(() => {
        }).catch((ex) => {
            console.log(ex)
        });
        }
        catch (error) {
            return error;
        }
    }

    // Close the task module on back navigation
    public onBack = () => {
        microsoftTeams.tasks.submitTask();
    }

    // Delete selected time range messages
    public onDelete = () => {
        let spanner = document.getElementsByClassName("prepareToDeleteLoader");
        spanner[0].classList.remove("hiddenLoader");
        
        this.deleteHistoricalMessages().then(() => {
           microsoftTeams.tasks.submitTask();
        });    
}

    // Show the selected time range in detail format on popup view 
    public dataRangeValue(): any {
        return (<div>{this.localize("From")} {this.state.startDate} {this.localize("To")} {this.state.endDate}</div>);
    }

    public render(): JSX.Element {

        return (
            (this.state.loader) ?
                <div className="Loader">
                    <Loader />
                </div>
                : <div className="deleteMessagesTaskModule">
                    <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                        <Flex className="scrollableContent">

                            <Flex.Item size="size.full">
                                <Flex column className="formContentContainer">
                                    <h3>{this.localize("DeleteMessagesPopup")}</h3>
                                    <h4>{this.localize("DataRange")}</h4>
                                    <div>{this.dataRangeValue()}</div>
                                </Flex>
                            </Flex.Item>
                        </Flex>
                        <Flex className="footerContainer" vAlign="end" hAlign="end">
                            <Flex className="buttonContainer" gap="gap.small">
                                <Text content={this.localize("DeleteConfirmationErrorMessage")} className="error-message" error size="medium" />
                                <Flex.Item push>
                                    <Loader id="prepareToDeleteLoader" className="hiddenLoader prepareToDeleteLoader" size="smallest" label={this.localize("PrepareToDeleteMessageLabel")} labelPosition="end" />
                                </Flex.Item>
                                <Flex.Item>
                                    <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                                </Flex.Item>
                                <Button content={this.localize("Delete")} id="deleteBtn" onClick={this.onDelete} primary />
                            </Flex>
                        </Flex>
                    </Flex>
                </div>
        );
    }
}

const messagesWithTranslation = withTranslation()(DeleteMessagesTaskModule);
export default (messagesWithTranslation);