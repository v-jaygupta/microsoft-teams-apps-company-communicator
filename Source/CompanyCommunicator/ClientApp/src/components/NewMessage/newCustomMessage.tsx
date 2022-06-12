// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { withTranslation } from "react-i18next";
import * as AdaptiveCards from "adaptivecards";
import { Button, Loader, Dropdown, Text, Flex, TextArea, RadioGroup, Checkbox } from '@fluentui/react-northstar'
import * as microsoftTeams from "@microsoft/teams-js";

import './newMessage.scss';
import './teamTheme.scss';
import { getDraftNotification, getTeams, createDraftNotification, updateDraftNotification, searchGroups, getGroups, verifyGroupAccess } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, 
} from '../AdaptiveCard/adaptiveCard';
import { ImageUtil } from '../../utility/imageutility';
import { TFunction } from "i18next";
import TimePicker, { LanguageDirection } from '../common/DateAndTimePicker/TimePicker';
import LocalizedDatePicker from '../common/DateAndTimePicker/LocalizedDatePicker';
import { formState, IDraftMessage, INewMessageProps } from './newMessage';

type dropdownItem = {
    key: string,
    header: string,
    content: string,
    image: string,
    team: {
        id: string,
    },
}


class NewCustomMessage extends React.Component<INewMessageProps, formState> {
    readonly localize: TFunction;
    private card: any;

    constructor(props: INewMessageProps) {
        super(props);
        this.localize = this.props.t;
        this.card = getInitAdaptiveCard(this.localize);
        this.setDefaultCard(this.card);

        const cardPlaceHolder = '{\r\n  \"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\",\r\n  \"type\": \"AdaptiveCard\",\r\n  \"version\": \"1.0\",\r\n  \"body\": [\r\n    {\r\n          \"type\": \"TextBlock\",\r\n          \"text\": \"This is an Adaptive Card\",\r\n          \"weight\": \"bolder\",\r\n          \"size\": \"medium\"\r\n    }\r\n  ]\r\n}\r\n';

        this.state = {
            title: "",
            summary: cardPlaceHolder,
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
            selectedRequestReadReceipt: false,
            selectedScheduledDateTime: new Date(),

            selectedFullWidth: false,
            selectedNotifyUser: false,
            selectedOnBehalfOf: false,
            selectedStageView: false,

            errorImageUrlMessage: "",
            errorButtonUrlMessage: "",
            usersList: "",
            messageType: 'CustomAC',
            templates: ['product announcement', 'conference', 'holidays'],
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
                    if (this.state.summary) {
                        adaptiveCard.parse(JSON.parse(this.state.summary));
                    }
                    else {
                        adaptiveCard.parse(this.state.card);
                    }
                    
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
        
        setCardTitle(card, titleAsString);
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
            else if (draftMessageDetail.usersList && draftMessageDetail.usersList.length > 0) {
                selectedRadioButton = "csv";
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
                selectedGroups: draftMessageDetail.groups,

                selectedRequestReadReceipt: draftMessageDetail.ack,                
                selectedScheduledDateTime: draftMessageDetail.scheduledDateTime !== null ? new Date(draftMessageDetail.scheduledDateTime) : draftMessageDetail.scheduledDateTime,
                selectedDelayDelivery: draftMessageDetail.scheduledDateTime !== null,
                
                selectedFullWidth: draftMessageDetail.fullWidth,
                selectedNotifyUser: draftMessageDetail.notifyUser,
                selectedInlineTranslation: draftMessageDetail.inlineTranslation,
                selectedOnBehalfOf: draftMessageDetail.onBehalfOf,
                selectedStageView: draftMessageDetail.stageView,
            });

            
            //setCardSummary(this.card, draftMessageDetail.summary);
            
            this.setState({
                title: draftMessageDetail.title,
                summary: draftMessageDetail.summary,

                btnLink: draftMessageDetail.buttonLink,
                imageLink: draftMessageDetail.imageLink,
                btnTitle: draftMessageDetail.buttonTitle,
                author: draftMessageDetail.author,
                allUsersOptionSelected: draftMessageDetail.allUsers,
                usersList: draftMessageDetail.usersList,
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
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                            <Flex className="scrollableContent">
                                <Flex.Item size="size.half">
                                    <Flex column className="formContentContainer">
                                        <Text content={this.localize("AdaptiveCardJSONPayload")} />
                                            <TextArea
                                                autoFocus
                                                placeholder={this.localize("AdaptiveCardJSONPayload")}
                                                value={this.state.summary}
                                                onChange={this.onSummaryChanged}
                                                id="summaryTextArea"
                                            fluid resize="both" variables={{ 'height': '400px' }} />
                                        <Text content={this.localize("AdaptiveCardJSONPayloadDescription") }/>                                        
                                        <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                    </Flex>
                                </Flex.Item>
                                <Flex.Item size="size.half">
                                    <div className="adaptiveCardContainer">
                                    </div>
                                </Flex.Item>
                            </Flex>

                            <Flex className="footerContainer" vAlign="end" hAlign="end">
                                <Flex className="buttonContainer">
                                    <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                </Flex>
                            </Flex>

                        </Flex>
                    </div>
                );
            }
            else if (this.state.page === "AudienceSelection") {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                            <Flex className="scrollableContent">
                                <Flex.Item size="size.half">
                                    <Flex column className="formContentContainer">
                                        <h3>{this.localize("SendHeadingText")}</h3>
                                        <RadioGroup
                                            className="radioBtns"
                                            checkedValue={this.state.selectedRadioBtn}
                                            onCheckedValueChange={this.onGroupSelected}
                                            vertical={true}
                                            items={[
                                                {
                                                    name: "teams",
                                                    key: "teams",
                                                    value: "teams",
                                                    label: this.localize("SendToGeneralChannel"),
                                                    children: (Component, { name, ...props }) => {
                                                        return (
                                                            <Flex key={name} column>
                                                                <Component {...props} />
                                                                <Dropdown
                                                                    hidden={!this.state.teamsOptionSelected}
                                                                    placeholder={this.localize("SendToGeneralChannelPlaceHolder")}
                                                                    search
                                                                    multiple
                                                                    items={this.getItems()}
                                                                    value={this.state.selectedTeams}
                                                                    onChange={this.onTeamsChange}
                                                                    noResultsMessage={this.localize("NoMatchMessage")}
                                                                />
                                                            </Flex>
                                                        )
                                                    },
                                                },
                                                {
                                                    name: "rosters",
                                                    key: "rosters",
                                                    value: "rosters",
                                                    label: this.localize("SendToRosters"),
                                                    children: (Component, { name, ...props }) => {
                                                        return (
                                                            <Flex key={name} column>
                                                                <Component {...props} />
                                                                <Dropdown
                                                                    hidden={!this.state.rostersOptionSelected}
                                                                    placeholder={this.localize("SendToRostersPlaceHolder")}
                                                                    search
                                                                    multiple
                                                                    items={this.getItems()}
                                                                    value={this.state.selectedRosters}
                                                                    onChange={this.onRostersChange}
                                                                    unstable_pinned={this.state.unstablePinned}
                                                                    noResultsMessage={this.localize("NoMatchMessage")}
                                                                />
                                                            </Flex>
                                                        )
                                                    },
                                                },
                                                {
                                                    name: "allUsers",
                                                    key: "allUsers",
                                                    value: "allUsers",
                                                    label: this.localize("SendToAllUsers"),
                                                    children: (Component, { name, ...props }) => {
                                                        return (
                                                            <Flex key={name} column>
                                                                <Component {...props} />
                                                                <div className={this.state.selectedRadioBtn === "allUsers" ? "" : "hide"}>
                                                                    <div className="noteText">
                                                                        <Text error content={this.localize("SendToAllUsersNote")} />
                                                                    </div>
                                                                </div>
                                                            </Flex>
                                                        )
                                                    },
                                                },
                                                {
                                                    name: "groups",
                                                    key: "groups",
                                                    value: "groups",
                                                    label: this.localize("SendToGroups"),
                                                    children: (Component, { name, ...props }) => {
                                                        return (
                                                            <Flex key={name} column>
                                                                <Component {...props} />
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
                                                                    onChange={this.onGroupsChange}
                                                                    noResultsMessage={this.state.noResultMessage}
                                                                    unstable_pinned={this.state.unstablePinned}
                                                                />
                                                                <div className={this.state.groupsOptionSelected && this.state.groupAccess ? "" : "hide"}>
                                                                    <div className="noteText">
                                                                        <Text error content={this.localize("SendToGroupsNote")} />
                                                                    </div>
                                                                </div>
                                                            </Flex>
                                                        )
                                                    },
                                                },
                                                {
                                                    name: "csv",
                                                    key: "csv",
                                                    value: "csv",
                                                    label: this.localize("SendToCSVUsers"),
                                                    children: (Component, { name, ...props }) => {
                                                        return (
                                                            <Flex key={name} column>
                                                                <Component {...props} />
                                                                <TextArea
                                                                    placeholder={this.localize("Users")}
                                                                    value={this.state.usersList}
                                                                    onChange={this.onUsersListChanged}
                                                                    id="csvUsers"
                                                                    disabled={this.state.selectedRadioBtn !== "csv"}
                                                                    fluid resize="both" variables={{ 'height': '150px' }} />
                                                                <div className={this.state.selectedRadioBtn === "csv" ? "" : "hide"}>
                                                                    <div className="noteText">
                                                                        <Text error content={this.localize("SupportedDelimiters")} />
                                                                    </div>
                                                                </div>
                                                            </Flex>
                                                        )
                                                    },
                                                },
                                            ]}
                                        >

                                        </RadioGroup>
                                        
                                    </Flex>
                                </Flex.Item>
                                <Flex.Item size="size.half">
                                    <div className="adaptiveCardContainer">
                                    </div>
                                </Flex.Item>
                            </Flex>
                            <Flex className="footerContainer" vAlign="end" hAlign="end">
                                <Flex className="buttonContainer">
                                    <Flex.Item push>
                                        <Button content={this.localize("Back")} disabled={this.isBackBtnDisabled()} onClick={this.onBack} secondary />
                                    </Flex.Item>
                                    <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                </Flex>
                            </Flex>
                        </Flex>
                    </div>
                );
            }
            else if (this.state.page === "AdditionalOptions") {                
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                            <Flex className="scrollableContent">
                                <Flex.Item size="size.half">
                                    <Flex column className="formContentContainer">
                                        <h3>{this.localize("SendOptions")}</h3>                                        
                                        <Checkbox label={this.localize("DelayDelivery")} checked={this.state.selectedDelayDelivery}
                                            onChange={this.onDelayDeliveryChanged} />

                                        <Flex gap="gap.smaller">
                                            <Flex.Item>
                                                <LocalizedDatePicker
                                                    screenWidth={500}
                                                    selectedDate={this.state.selectedScheduledDateTime}
                                                    minDate={new Date()}
                                                    onDateSelect={this.onDeliveryDateChanged}
                                                    disableSelection={!this.state.selectedDelayDelivery} theme={""}
                                                />
                                            </Flex.Item>
                                            <Flex.Item>
                                                <TimePicker
                                                    hours={this.state.selectedScheduledDateTime === undefined ? 0 : new Date(this.state.selectedScheduledDateTime).getHours()}
                                                    minutes={this.state.selectedScheduledDateTime === undefined ? 0 : new Date(this.state.selectedScheduledDateTime).getMinutes()}
                                                    isDisabled={!this.state.selectedDelayDelivery}
                                                    onPickerClose={this.onDeliveryTimeChange}
                                                    dir={LanguageDirection.Ltr} />
                                            </Flex.Item>
                                        </Flex>
                                        
                                        <Checkbox label={this.localize("NotifyUser")} checked={this.state.selectedNotifyUser} onChange={this.onNotifyUserChanged} />
                                        
                                    </Flex>
                                </Flex.Item>
                                <Flex.Item size="size.half">
                                    <div className="adaptiveCardContainer">
                                    </div>
                                </Flex.Item>
                            </Flex>

                            <Flex className="footerContainer" vAlign="end" hAlign="end">
                                <Flex className="buttonContainer" gap="gap.small">
                                    <Flex.Item push>
                                        <Loader id="sendingLoader" className="hiddenLoader sendingLoader" size="smallest" label={this.localize("PreparingMessageLabel")} labelPosition="end" />
                                    </Flex.Item>
                                    <Flex.Item push>
                                        <Button content={this.localize("Back")} disabled={this.isBackBtnDisabled()} onClick={this.onBack} secondary />
                                    </Flex.Item>
                                    <Button content={this.localize("SaveAsDraft")} disabled={this.isSaveBtnDisabled()} id="saveBtn" onClick={this.onSave} primary />
                                </Flex>
                            </Flex>
                        </Flex>
                    </div>
                );
            }
            else {
                return (<div>Error</div>);
            }
        }
    }

    private onRequestReadReceiptChanged = (event: any, data: any) => {
        this.setState({
            selectedRequestReadReceipt: data.checked,
        });
    }

    private onDelayDeliveryChanged = (event: any, data: any) => {
        this.setState({
            selectedDelayDelivery: data.checked,
        })
    }

    private onInlineTranslationChanged = (event: any, data: any) => {
        this.setState({
            selectedInlineTranslation: data.checked,
        })
    }

    private onFullWidthChanged = (event: any, data: any) => {
        this.setState({
            selectedFullWidth: data.checked,
        })
    }

    private onNotifyUserChanged = (event: any, data: any) => {
        this.setState({
            selectedNotifyUser: data.checked,
        })
    }

    private onBehalfOfChanged = (event: any, data: any) => {
        this.setState({
            selectedOnBehalfOf: data.checked,
        })
    }

    private onStageViewChanged = (event: any, data: any) => {
        this.setState({
            selectedStageView: data.checked,
        })
    }

    private onDeliveryTimeChange = (hours: number, min: number) => {
        var date = this.state.selectedScheduledDateTime === undefined ?
            new Date(new Date().setHours(hours, min)) : new Date(new Date(this.state.selectedScheduledDateTime).setHours(hours, min));
        this.setState({ selectedScheduledDateTime: date });
    }

    private onDeliveryDateChanged = (date: Date) => {
        console.log(date);
        this.setState({ selectedScheduledDateTime: date });
    }

    private onUsersListChanged = (event: any): void => {
        let list = event.target.value;
        list = list.replace(/(?:\\[rn]|[\r\n]+)+/g, ",");
        this.setState({
            usersList: list,
        });
    }


    private onGroupSelected = (event: any, data: any) => {
        this.setState({
            selectedRadioBtn: data.value,
            teamsOptionSelected: data.value === 'teams',
            rostersOptionSelected: data.value === 'rosters',
            groupsOptionSelected: data.value === 'groups',
            allUsersOptionSelected: data.value === 'allUsers',
            selectedTeams: data.value === 'teams' ? this.state.selectedTeams : [],
            selectedTeamsNum: data.value === 'teams' ? this.state.selectedTeamsNum : 0,
            selectedRosters: data.value === 'rosters' ? this.state.selectedRosters : [],
            selectedRostersNum: data.value === 'rosters' ? this.state.selectedRostersNum : 0,
            selectedGroups: data.value === 'groups' ? this.state.selectedGroups : [],
            selectedGroupsNum: data.value === 'groups' ? this.state.selectedGroupsNum : 0,
        });
    }

    private isSaveBtnDisabled = () => {
        const customUsers = (this.state.usersList !== null && this.state.usersList.length > 0);
        if (customUsers && (this.state.selectedRadioBtn === "csv")) {
            return false;
        }
        const teamsSelectionIsValid = (this.state.teamsOptionSelected && (this.state.selectedTeamsNum !== 0)) || (!this.state.teamsOptionSelected);
        const rostersSelectionIsValid = (this.state.rostersOptionSelected && (this.state.selectedRostersNum !== 0)) || (!this.state.rostersOptionSelected);
        const groupsSelectionIsValid = (this.state.groupsOptionSelected && (this.state.selectedGroupsNum !== 0)) || (!this.state.groupsOptionSelected);        
        const nothingSelected = (!this.state.teamsOptionSelected) && (!this.state.rostersOptionSelected) && (!this.state.groupsOptionSelected) && (!this.state.allUsersOptionSelected);
        return (!teamsSelectionIsValid || !rostersSelectionIsValid || !groupsSelectionIsValid || nothingSelected)
    }

    private isNextBtnDisabled = () => {
        const current = this.state.page;
        if (current === "AudienceSelection") {
            return this.isSaveBtnDisabled();
        }
        else {
            const title = this.state.title;
            return !(title);
        }
    }

    private isBackBtnDisabled = () => {
        const current = this.state.page;
        if (current === "AudienceSelection") {
            return this.isSaveBtnDisabled();
        }
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
        if (itemsData.value.length > NewCustomMessage.MAX_SELECTED_TEAMS_NUM) return;
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
        if (itemsData.value.length > NewCustomMessage.MAX_SELECTED_TEAMS_NUM) return;
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
                noResultMessage: this.localize("NoMatchMessage"),
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
                    noResultMessage: this.localize("NoMatchMessage")
                });
            }
            catch (error) {
                return error;
            }
        }
    }

    private onSave = () => {
        let spanner = document.getElementsByClassName("sendingLoader");
        spanner[0].classList.remove("hiddenLoader");

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
            allUsers: this.state.allUsersOptionSelected,
            ack: this.state.selectedRequestReadReceipt,
            scheduledDateTime: this.state.selectedDelayDelivery ? this.state.selectedScheduledDateTime : undefined,

            fullWidth: this.state.selectedFullWidth,
            notifyUser: this.state.selectedNotifyUser,
            inlineTranslation: this.state.selectedInlineTranslation,
            onBehalfOf: this.state.selectedOnBehalfOf,
            stageView: this.state.selectedStageView,
            messageType: this.state.messageType,
            usersList: this.state.selectedRadioBtn === "csv" ? this.state.usersList : "",
        };

        if (this.state.exists) {
            this.editDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        } else {
            this.postDraftMessage(draftMessage).then(() => {
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

    private postDraftMessage = async (draftMessage: IDraftMessage) => {
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
        const current = this.state.page;
        let next: string = (current === "CardCreation") ? "AudienceSelection" : "AdditionalOptions";

        this.setState({
            page: next
        }, () => {
            this.updateCard();
        });
    }

    private onBack = (event: any) => {
        const current = this.state.page;
        let back: string = (current === "AdditionalOptions") ? "AudienceSelection" : "CardCreation";

        this.setState({
            page: back
        }, () => {
            this.updateCard();
        });
    }

    

    

    private onSummaryChanged = (event: any) => {
        /*let showDefaultCard = (!this.state.title && !this.state.imageLink && !event.target.value && !this.state.author && !this.state.btnTitle && !this.state.btnLink);*/
        
        //setFullCardContent(this.card, event.target.value);
        
        this.setState({
            summary: event.target.value,
            title: this.localize("CustomAdaptiveCard")
        }, () => {
            console.log(this.state);
            this.updateCard();
        });
    }

    

    

    private updateCard = () => {
        const adaptiveCard = new AdaptiveCards.AdaptiveCard();
        if (this.state.summary) {
            adaptiveCard.parse(JSON.parse(this.state.summary));
        }
        else {
            adaptiveCard.parse(this.state.card);
        }
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
}

const newCustomMessageWithTranslation = withTranslation()(NewCustomMessage);
export default newCustomMessageWithTranslation;
