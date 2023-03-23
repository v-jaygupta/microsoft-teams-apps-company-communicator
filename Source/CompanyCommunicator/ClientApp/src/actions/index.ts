// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { formatDate } from '../i18n';
import { getSentNotifications, getDraftNotifications, deleteHistoricalData, getDeleteMessagesData } from '../apis/messageListApi';

type Notification = {
    createdDateTime: string,
    failed: number,
    id: string,
    isCompleted: boolean,
    sentDate: string,
    sendingStartedDate: string,
    sendingDuration: string,
    succeeded: number,
    throttled: number,
    title: string,
    totalMessageCount: number,
    createdBy: string,
}

type cleanUpHistory = {
    timestamp:string,
    status: string,
    selectedDateRange: string,
    recordsDeleted: number,
    deletedBy: string
}

export const selectMessage = (message: any) => {
    return {
        type: 'MESSAGE_SELECTED',
        payload: message
    };
};

export const getMessagesList = () => async (dispatch: any) => {
    const response = await getSentNotifications();
    const notificationList: Notification[] = response.data;
    notificationList.forEach(notification => {
        notification.sendingStartedDate = formatDate(notification.sendingStartedDate);
        notification.sentDate = formatDate(notification.sentDate);
    });
    dispatch({ type: 'FETCH_MESSAGES', payload: notificationList });
};

export const getDraftMessagesList = () => async (dispatch: any) => {
    const response = await getDraftNotifications();
    dispatch({ type: 'FETCH_DRAFTMESSAGES', payload: response.data });
};

// Get deleted messages list
export const getDeleteMessagesList = () => async (dispatch: any) => {
    const response = await getDeleteMessagesData();
    const cleanUpHistoryList: cleanUpHistory[] = response.data;
    dispatch({ type: 'FETCH_DELETEMESSAGES', payload: cleanUpHistoryList });
};