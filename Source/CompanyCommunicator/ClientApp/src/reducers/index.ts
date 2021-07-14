// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { combineReducers } from "redux";

export const selectedMessageReducer = (selectedMessage = null, action: { type: string; payload: any; }) => {
    if (action.type === 'MESSAGE_SELECTED') {
        return action.payload;
    }
    return selectedMessage;
}

export const messagesListReducer = (messages = [], action: { type: string; payload: any; }) => {
    if (action.type === 'FETCH_MESSAGES') {
        return action.payload
    }
    return messages;
};

export const draftmessagesListReducer = (draftMessages = [], action: { type: string; payload: any; }) => {
    if (action.type === 'FETCH_DRAFTMESSAGES') {
        return action.payload
    }
    return draftMessages;
};

export const scheduledmessagesListReducer = (scheduledMessages = [], action: { type: string; payload: any; }) => {
    if (action.type === 'FETCH_SCHEDULEDMESSAGES') {
        return action.payload
    }
    return scheduledMessages;
};

export default combineReducers({
    messagesList: messagesListReducer,
    draftMessagesList: draftmessagesListReducer,
    scheduledMessagesList: scheduledmessagesListReducer,
    selectedMessage: selectedMessageReducer,
});