import { SPHttpClientResponse, ISPHttpClientOptions, SPHttpClient } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import Constants from './Constants';

export default class CommonUtility {

    //Get Current User canManageWeb permission
    public isUserHasCanManageWebPermission(userPermissionValue: any) {
        let permission = new SPPermission(userPermissionValue);
        return permission.hasPermission(SPPermission.manageWeb);
    }

    //Get list details with all its inbuilt fields.
    public getData(url: string): Promise<any> {
        return fetch(url, { credentials: 'same-origin', headers: { Accept: 'application/json;odata=verbose', "Content-Type": 'application/json;odata=verbose' } }).then((response) => {
            return response.json();
        }, (errorFail) => {
            console.log("error");
        }).then((responseJSON) => {
            return responseJSON.d.results;//.AllItems;
        }).catch((error) => {
            return null;
        });
    }

    //will get list entity type
    public getListItemEntityTypeName(siteAbsoluteUrl: string, listname: string): Promise<string> {
        var url = `${siteAbsoluteUrl}/_api/web/lists/getbytitle('${listname}')?$select=ListItemEntityTypeFullName`;
        return fetch(url, { credentials: 'same-origin', headers: { Accept: 'application/json;odata=nometadata', "Content-Type": 'application/json;odata=verbose', 'Access-Control-Allow-Origin': '*' } }).then((response) => {
            return response.json();
        }, (errorFail) => {
            console.log("error");
        }).then((responseJSON) => {
            return responseJSON.ListItemEntityTypeFullName;
        }).catch((response: SPHttpClientResponse) => {
            return null;
        });
    }

    //will fetch request digest value
    public getValues(site): any {
        try {
            var url = site + '/_api/contextinfo';
            return fetch(url, {
                method: "POST",
                headers: { Accept: "application/json;odata=verbose", "Content-Type": 'application/json;odata=verbose', 'Access-Control-Allow-Origin': '*' },
                credentials: "same-origin"
            }).then((response) => {
                return response.json();
            });
        } catch (error) {
            console.log("getValues: " + error);
        }
    }

    //will retrieve data based on url
    public getRequest(url: string): any {
        try {
            return fetch(url, {
                headers: { Accept: 'application/json;odata=verbose', "Content-Type": 'application/json;odata=verbose', 'Access-Control-Allow-Origin': '*' },
                credentials: "same-origin"
            }).then((response) => {
                if (response.status >= 200 && response.status < 400) {
                    return response.json();
                }
                else {
                    return response.json();
                }
            }).catch(error =>
                console.error("getRequest: " + error));
        } catch (error) {
            console.log("getRequest: " + error);
        }
    }

    //will send post request to add/update list data
    public postRequest(url: string, postBody, xMethod, requestDigestValue): any {
        try {
            //GET FormDigestValue
            return fetch(url, {
                headers: {
                    Accept: 'application/json;odata=verbose',
                    "Content-Type": 'application/json;odata=verbose',
                    "X-RequestDigest": requestDigestValue,
                    "X-Http-Method": xMethod,
                    'IF-MATCH': '*'
                },
                method: 'POST',
                body: postBody,
                credentials: "same-origin"
            }).then((response) => {
                if (response.status <= 204 && response.status >= 200) {
                    //resolve(response);
                    if (response.status == 204) {
                        return 'success';
                    } else {
                        return response.json();
                    }
                }
                else {
                    return response.json();
                }
            }, (err) => {
                console.log(err);
            }).catch(error =>
                console.error("postRequest: " + error));

        } catch (error) {
            console.log("postRequest: " + error);
        }
    }

    //will check if folder exists of job form ID and will call method to create a folder
    public checkFolderExists(url, siteAbsoluteURL, requestDigestValue) {
        return fetch(url, { credentials: 'same-origin', headers: { Accept: 'application/json;odata=verbose', "Content-Type": 'application/json;odata=verbose', 'Access-Control-Allow-Origin': '*' } }).then((response) => {
            return response.json();
        }, (errorFail) => {
            console.log("error");
        }).then((responseJSON) => {
            if (responseJSON.d != undefined) {
                return responseJSON.d.results;//.AllItems;
            }
            else {
                let urlToCreateFolder = `${siteAbsoluteURL}/_api/web/lists/getbytitle('${Constants.ccpDocumentsListTitle}')/RootFolder/folders/add(url='3')`;
                this.postRequestWithoutBody(urlToCreateFolder, requestDigestValue, '').then(
                    (results: any): void => {
                        console.log(results);
                        return results;
                    }).catch((error) => {
                        console.log(error);
                        return error;
                    });
            }
        }).catch((error) => {
            console.log(error);
        });
    }

    //will create folder of job form id
    public postRequestWithoutBody(urlToCreateFolder, requestDigestValue, body) {
        try {
            //GET FormDigestValue
            return fetch(urlToCreateFolder, {
                headers: {
                    Accept: 'application/json;odata=verbose',
                    "Content-Type": 'application/json;odata=verbose',
                    "X-RequestDigest": requestDigestValue
                },
                method: 'POST',
                credentials: "same-origin",
                body: body
            }).then((response) => {
                return response.json();
            }, (err) => {
                console.log(err);
            }).catch(error =>
                console.error("postRequest: " + error));

        } catch (error) {
            console.log("postRequest: " + error);
        }
    }

    //will upload files to document library
    public async uploadFiles(uploadFileObj, fileName, context, siteurl, jobId) {
        try {
            if (uploadFileObj != '') {
                var file = uploadFileObj;
                if (file != undefined || file != null) {
                    let spOpts: ISPHttpClientOptions = {
                        headers: {
                            "Accept": "application/json",
                            "Content-Type": "application/json"
                        },
                        body: file,
                        credentials: "same-origin"
                    };
                    let url = `${siteurl}/_api/Web/Lists/getByTitle('${Constants.ccpDocumentsListTitle}')/RootFolder/folders('${jobId}')/Files/Add(url='${fileName}', overwrite=true)`;

                    //POST call
                    return context.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts).then(async (response: SPHttpClientResponse) => {
                        return response.json();
                    }).then(async (responseJSON: JSON) => {
                        return true;
                    });
                }
            } else {
                return true;
            }
        }
        catch (error) {
            console.log('uploadFiles : ' + error);
        }
    }

    //will send post request to delete list data
    public deleteRequest(url: string, xMethod, requestDigestValue): any {
        try {
            //GET FormDigestValue
            return fetch(url, {
                headers: {
                    Accept: 'application/json;odata=verbose',
                    "Content-Type": 'application/json;odata=verbose',
                    "X-RequestDigest": requestDigestValue,
                    "X-Http-Method": xMethod,
                    'IF-MATCH': '*'
                },
                method: 'POST',
                credentials: "same-origin"
            }).then((response) => {
                if (response.status <= 204 && response.status >= 200) {
                    //resolve(response);
                    if (response.status == 204 || xMethod == 'DELETE') {
                        return 'success';
                    } else {
                        return response.json();
                    }
                }
                else {
                    return response.json();
                }
            }, (err) => {
                console.log(err);
            }).catch(error =>
                console.error("postRequest: " + error));

        } catch (error) {
            console.log("postRequest: " + error);
        }
    }

    //will set standarad date format
    public setStandardDateFormat(date: Date) {
        return (date.getDate() >= 10 ? date.getDate() : ('0' + date.getDate())) + '/' + (((date.getMonth() + 1) >= 10 ? (date.getMonth() + 1) : ('0' + (date.getMonth() + 1))) + '/' + date.getFullYear());
    }

    //will set am pm time format
    public setAMPMTimeFormat(d: Date) {
        //var d = new Date(dateInput);
        var ampm = (d.getHours() >= 12) ? "PM" : "AM";
        var hours = (d.getHours() >= 12) ? d.getHours() - 12 : d.getHours();

        return (hours >= 10 ? hours : '0' + hours) + ':' + (d.getMinutes() >= 10 ? d.getMinutes() : '0' + d.getMinutes()) + ' ' + ampm;
    }

    //will get and set user group for logged in user based on the details
    public async getAndSetCurrentUserGroupsOnState(currentUserId: string, absoluteUrl, context, currentSite, formID?) {
        var userGroups: any = [];
        let currentUserGroups = "";

        //NOTE: Check current user has Admin permission.
        if (context.pageContext.legacyPageContext.isSiteAdmin) {
            currentUserGroups = Constants.userGroups[0].title + ",";
        }
        var userGrpAPI = absoluteUrl + "/_api/web/siteusers/getById(" + currentUserId + ")/groups";
        userGroups = await this.getData(userGrpAPI);
        userGroups.map(async (grp) => {
            if (grp.Title == Constants.userGroups[6].title) {
                currentUserGroups += Constants.userGroups[6].title + ",";
            }
            if (grp.Title == Constants.userGroups[2].title) {
                currentUserGroups += Constants.userGroups[2].title + ",";
            }
            if (grp.Title == Constants.userGroups[3].title) {
                currentUserGroups += Constants.userGroups[3].title + ",";
            }
            if (grp.Title == Constants.userGroups[4].title) {
                currentUserGroups += Constants.userGroups[4].title + ",";
            }


        });
        var url = "";
        if (formID != undefined) {
            url = `${currentSite}/_api/web/lists/getbytitle('${Constants.CCPJobListTitle}')/items?$select=ID, Job_x0020_Status, Is_x0020_job_x0020_ready_x0020_f, ManagerNameId&$orderby=Title&$filter=ID eq ${formID}`;
        }
        else {
            url = `${currentSite}/_api/web/lists/getbytitle('${Constants.CCPJobListTitle}')/items?$select=ID, Job_x0020_Status, Is_x0020_job_x0020_ready_x0020_f, ManagerNameId&$orderby=Title`;

        }

        await this.getData(url).then(
            (results: any) => {
                results.map((obj, key) => {
                    if (obj.Is_x0020_job_x0020_ready_x0020_f && obj.Job_x0020_Status == "Awaiting manager approval" && obj.ManagerNameId == currentUserId) {
                        currentUserGroups += Constants.userGroups[1].title + ",";
                    }
                });
            }
        );
        if (currentUserGroups == "") {
            currentUserGroups += Constants.userGroups[5].title;
        }
        await localStorage.setItem('Current User Groups', currentUserGroups);
        return currentUserGroups;
    }

    //will perform post call for insertion of data to list
    public insertdata(siteAbsoluteUrl: string, listname: string, requestdata): Promise<number> {
        var url = `${siteAbsoluteUrl}/_api/web/lists/getbytitle('${listname}')/items`;
        return new Promise<number>((resolve, reject) => {
            try {
                this.getValues(siteAbsoluteUrl).then((token) => {
                    fetch(url,
                        {
                            method: "POST",
                            credentials: 'same-origin',
                            headers: {
                                Accept: 'application/json',
                                "Content-Type": "application/json;odata=verbose",
                                "X-RequestDigest": token.d.GetContextWebInformation.FormDigestValue
                            },
                            body: requestdata,
                        }).then((response) => {
                            return response.json();
                        }).then((response) => {

                            resolve(response.Id);
                            return response.Id;
                        }).catch((error) => {
                            reject(error);
                        });
                }, (error) => {
                    console.log(error);
                    reject(error);
                });
            }
            catch (e) {
                console.log(e);
                reject(e);
            }
        });
    }

    //will perform post call for updation of data to list
    public updatedata(siteAbsoluteUrl: string, listname: string, requestdata, id: number) {
        var url = `${siteAbsoluteUrl}/_api/web/lists/getbytitle('${listname}')/items(${id})`;
        return new Promise<any>((resolve, reject) => {
            try {
                this.getValues(siteAbsoluteUrl).then((token) => {
                    fetch(url,
                        {
                            method: "POST",
                            credentials: 'same-origin',
                            headers: {
                                Accept: 'application/json',
                                "Content-Type": "application/json;odata=verbose",
                                "X-RequestDigest": token.d.GetContextWebInformation.FormDigestValue,
                                'IF-MATCH': '*',
                                'X-HTTP-Method': 'MERGE'
                            },
                            body: requestdata,
                        })
                        .then((response) => {

                            resolve(response);
                            return response;
                        })
                        .catch((error) => {
                            reject(error);
                        });
                }, (error) => {
                    console.log(error);
                    reject(error);
                });
            }
            catch (e) {
                console.log(e);
                reject(e);
            }
        });
    }

    //Get the user id based on user login name
    public getUserId(url: string): Promise<any> {
        return fetch(url, { credentials: 'same-origin', headers: { Accept: 'application/json;odata=verbose', "Content-Type": 'application/json;odata=verbose', 'Access-Control-Allow-Origin': '*' } }).then((response) => {
            return response.json();
        }, (errorFail) => {
            console.log("error");
        }).then((responseJSON) => {
            return responseJSON.d;//.AllItems;
        }).catch((response: SPHttpClientResponse) => {
            return null;
        });
    }
}