import axios from './axiosJWTDecorator';
import { getBaseUrl } from '../configVariables';
import { AxiosRequestConfig } from 'axios';

let baseAxiosUrl = getBaseUrl() + '/api';

export const getSentNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications";
    return await axios.get(url);
}

export const getDraftNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.get(url);
}

export const verifyGroupAccess = async (): Promise<any> => {
    let url = baseAxiosUrl + "/groupdata/verifyaccess";
    return await axios.get(url, false);
}

export const getGroups = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/groupdata/" + id;
    return await axios.get(url);
}

export const searchGroups = async (query: string): Promise<any> => {
    let url = baseAxiosUrl + "/groupdata/search/" + query;
    return await axios.get(url);
}

export const exportNotification = async (id: string): Promise<any> => {
    let url = baseAxiosUrl + "/exportnotification/" + id;
    return await axios.put(url,null,false);
}

export const getSentNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications/" + id;
    return await axios.get(url);
}

export const getDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/" + id;
    return await axios.get(url);
}


export const deleteDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/" + id;
    return await axios.delete(url);
}

export const duplicateDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/duplicates/" + id;
    return await axios.post(url);
}

export const sendDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications";
    return await axios.post(url, payload);
}

export const updateDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.put(url, payload);
}

export const createDraftNotification = async (payload: any): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    /*const config: AxiosRequestConfig = {
        headers: {
            'Content-Type': 'multipart/form-data'
        }
    }*/
    return await axios.post(url, payload);
}

export const getTeams = async (): Promise<any> => {
    let url = baseAxiosUrl + "/teamdata";
    return await axios.get(url);
}

export const getConsentSummaries = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/consentSummaries/" + id;
    return await axios.get(url);
}

export const sendPreview = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/previews";
    return await axios.post(url, payload);
}

export const getAuthenticationConsentMetadata = async (windowLocationOriginDomain: string, login_hint: string): Promise<any> => {
    let url = `${baseAxiosUrl}/authenticationMetadata/consentUrl?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${login_hint}`;
    return await axios.get(url, undefined, false);
}
