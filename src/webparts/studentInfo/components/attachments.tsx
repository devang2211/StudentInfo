// import * as React from 'react';
// // import styles from './CommunityWebpart.module.scss';
// // import { ICommunityWebpartProps } from './ICommunityWebpartProps';
// import { escape } from '@microsoft/sp-lodash-subset';
// // import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
// import { SPComponentLoader } from '@microsoft/sp-loader';
// // import { CommunityWebpartState, IDiscussionThread, ITopicComments } from "./CommunityWebpartState";
// // import { SPGet, SPPost, SPUpdate, SPDelete } from "../../SPOHelper";
// import { Link, Spinner, SpinnerSize, Dropdown, PrimaryButton, DefaultButton, IconButton, IIconProps, TextField } from 'office-ui-fabric-react';
// import { Modal } from 'office-ui-fabric-react/lib/Modal';
// import { SPHttpClientResponse, ISPHttpClientOptions, SPHttpClient } from '@microsoft/sp-http';
// // import { Web, ConsoleListener } from "sp-pnp-js/lib/pnp";
// import { isEmpty } from '@microsoft/sp-lodash-subset';
// import { Icon } from 'office-ui-fabric-react/lib/Icon';
// import { MSGraphClient } from '@microsoft/sp-http';
// // import AutoComplateText from './AutoComplateText';

// var siteURL = "";
// const listTitle = "Community Photos";
// var rootUrl = window.location.origin;
// var promise;
// var Files = [];
// const listEntityTypeName = "SP.Data.Community_x0020_PhotosItem";
// const Community_User_Replies_List = "Community User Replies";
// const Community_Like_List = "Community Likes";
// var currentLoginUserInfo;
// var basesiteurl = window.location.origin;
// var basehuburl = basesiteurl + "/sites/contentTypeHub";
// var pinsymbol = basehuburl + "/Style%20Library/Community%20WebPart/images/pin.png";
// var filterCommentData = [];
// var totalLike = [];
// var countLike = 0;
// const currentDate = new Date();
// const currentYear = currentDate.getUTCFullYear();
// const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
// export default class CommunityWebpart extends React.Component<ICommunityWebpartProps, CommunityWebpartState> {
//     public web: Web = new Web(this.props.customContext.pageContext.web.absoluteUrl);
//     public imgList = ['gif', 'png', 'jpg', 'jpeg'];
//     public constructor(props: ICommunityWebpartProps, state: CommunityWebpartState) {
//         super(props);
//         siteURL = props.siteAbsoluteUrl;
//         this.state = {
//             CommunityPhotosValues: [],
//             CommunityLikeValues: [],
//             UserDetails: [],
//             AllComments: [],
//             error: null,
//             isLoaded: false,
//             showPopup: false,
//             popUpImageSrc: "",
//             postDescription: "",
//             showFileErrorMessage: false,
//             fileToSend: "",
//             postDscriptionErrorMessage: "",
//             fileErrorMessage: "",
//             unFilteredData: [],
//             dataValues: [],
//             discussionThreadValue: {} as IDiscussionThread,
//             topicComments: [] as ITopicComments[],
//             queryStringID: "",
//             newInsertedId: "",
//             TopicsThreadCount: 0,
//             PostImg: [],
//             isLike: false,
//             IsCommentClick: false,
//             UserShownInModel: "",
//             showDetailPopup: false,
//             IsCommentUserID: "",
//             text: "",
//             suggestions: [],
//             IsReplyDone: false,
//             totalComments: 0,
//             currentItemID: ''
//         };

//         SPComponentLoader.loadCss(window.location.origin + "/sites/contentTypeHub/Style%20Library/Community%20WebPart/css/CommunityWebPart.css");
//     }

//     public async getusr(graph: ICommunityWebpartProps) {
//         await graph.contextGraph.getClient()
//             .then(async (client: MSGraphClient): Promise<void> => {
//                 // get information about the current user from the Microsoft Graph
//                 await client
//                     .api('users').select("displayName,givenName,jobTitle,mail,mobilePhone,officeLocation,surname,businessPhones,assignedLicenses,assignedPlans,userPrincipalName,department")
//                     .filter("userType eq 'Member' AND accountEnabled eq true")
//                     .top(999)
//                     .version("v1.0")
//                     .get(async (error, response: any, rawResponse?: any) => {
//                         let ress = await response.value;
//                         ress = ress.sort(this.compareValues('displayName'));
//                         this.setState({ UserDetails: ress });
//                     });
//             });
//     }

//     public compareValues(key, order = 'asc') {
//         return function innerSort(a, b) {
//             if (!a.hasOwnProperty(key) || !b.hasOwnProperty(key)) {
//                 // property doesn't exist on either object
//                 return 0;
//             }

//             const varA = (typeof a[key] === 'string')
//                 ? a[key].toUpperCase() : a[key];
//             const varB = (typeof b[key] === 'string')
//                 ? b[key].toUpperCase() : b[key];

//             let comparison = 0;
//             if (varA > varB) {
//                 comparison = 1;
//             } else if (varA < varB) {
//                 comparison = -1;
//             }
//             return (
//                 (order === 'desc') ? (comparison * -1) : comparison
//             );
//         };
//     }

//     public componentDidMount() {
//         currentLoginUserInfo = this.props.customContext.pageContext.legacyPageContext;
//         this.getusr(this.props);
//         this.getCommunityPhotoDetails();
//         this.GetDataFunction(1);
//     }

//     public GetDataFunction(UserId) {
//         this.getDicussionComments(UserId);
//     }

//     public getDicussionComments(UserId) {
//         var Topiccomments: ITopicComments[] = [];
//         var AllComments = [];
//         promise = new Promise<any>((resolve, reject) => {
//             SPGet(
//                 siteURL +
//                 `/_api/web/lists/getByTitle('` +
//                 Community_User_Replies_List +
//                 `')/items?$select=Reply,CommunityUserId,Is_x0020_Parent,Thread_x0020_Parent,ID,Created,Author/Id,Author/Title,Author/EMail,AttachmentFiles&$expand=Author/Id&$expand=AttachmentFiles/ServerRelativeUrl`
//             ).then((r) => resolve(r));
//         });
//         promise
//             .then((r) => {
//                 {
//                     if (r.value.length > 0) {
//                         this.setState({
//                             AllComments: r.value,
//                         });
//                     }
//                 }
//             })
//             .catch((reject) => {

//             });
//     }

//     public getDicussionCommentsAfterReply(UserId, SelectedUser) {
//         var Topiccomments: ITopicComments[] = [];
//         var AllComments = [];
//         promise = new Promise<any>((resolve, reject) => {
//             SPGet(
//                 siteURL +
//                 `/_api/web/lists/getByTitle('` +
//                 Community_User_Replies_List +
//                 `')/items?$select=Reply,CommunityUserId,Is_x0020_Parent,Thread_x0020_Parent,ID,Created,Author/Id,Author/Title,Author/EMail,AttachmentFiles&$expand=Author/Id&$expand=AttachmentFiles/ServerRelativeUrl`
//             ).then((r) => resolve(r));
//         });
//         promise
//             .then((r) => {
//                 {
//                     if (r.value.length > 0) {
//                         this.setState({
//                             AllComments: r.value,
//                             IsReplyDone: false,
//                         });
//                         this.handleCommentClick(SelectedUser);
//                     }
//                 }
//             })
//             .catch((reject) => {

//             });
//     }

//     /**
//    * 
//    * @param ProductIndex 
//    * Togglepopup model click event
//    */
//     public togglePopup() {
//         this.setState({
//             showPopup: !this.state.showPopup,
//             postDescription: "",
//             postDscriptionErrorMessage: "",
//             fileToSend: "",
//             showFileErrorMessage: false,
//             fileErrorMessage: "",
//         });
//     }

//     public oneElementValueChange(statename, e, errorMessageStateName, entittyName) {
//         var state = this.state;
//         state[statename] = e.target.value;
//         state[errorMessageStateName] = "";
//         this.setState(state);
//     }

//     public async onFileUploadChange(e) {
//         var file = e.target.files[0];
//         Files = [];
//         if (file.name.split(".")[file.name.split(".").length - 1].toLowerCase() != "gif" && file.name.split(".")[file.name.split(".").length - 1].toLowerCase() != "png" && file.name.split(".")[file.name.split(".").length - 1].toLowerCase() != "jpg" && file.name.split(".")[file.name.split(".").length - 1].toLowerCase() != "jpeg" && file.name.split(".")[file.name.split(".").length - 1].toLowerCase() != "bmp") {
//             await this.setState({ showFileErrorMessage: false, fileErrorMessage: "" });
//         }
//         else {
//             await this.setState({ showFileErrorMessage: true, fileToSend: file.name, fileErrorMessage: "" });
//             Files.push(file);

//         }
//     }


//     public validatePostsDetails() {
//         var isValid = true;
//         if (this.state.postDescription.trim() == "") {
//             isValid = false;
//             this.setState({ postDscriptionErrorMessage: "Please enter Description" });
//         }
//         else {
//             this.setState({ postDscriptionErrorMessage: "" });
//         }
//         if (Files.length == 0) {
//             isValid = false;
//             this.setState({ fileErrorMessage: "Please select the file" });
//         }
//         else {
//             this.setState({ fileErrorMessage: "" });
//         }

//         if (isValid) {
//             this.startloader();
//             this.addPostDetailsToList();
//         }
//     }

//     public addPostDetailsToList() {
//         this.getValues(this.props.siteAbsoluteUrl)
//             .then(async (token) => {
//                 Files.forEach(async (object, index) => {
//                     let arrayBuffer = await this.getFileBuffer(object);
//                     await this.uploadFiles(object, object.name, this.props.customContext, listTitle, this.props.siteAbsoluteUrl, token.d.GetContextWebInformation.FormDigestValue);
//                     var unFilteredKdcData = await this.getCommunityPhotoDetails();
//                     this.setState({ unFilteredData: unFilteredKdcData, dataValues: unFilteredKdcData, showPopup: false });
//                     this.stoploader();
//                 });
//             });
//     }

//     private getFileBuffer(file) {
//         return new Promise<any>((resolve, reject) => {
//             var reader = new FileReader();
//             reader.onloadend = ((e: Event) => {
//                 resolve(reader.result);
//             });
//             reader.onerror = ((e: Event) => {
//                 reject(reader.error);
//             });
//             reader.readAsArrayBuffer(file);
//         });
//     }

//     public getValues(site): any {
//         try {
//             var url = site + '/_api/contextinfo';
//             return fetch(url, {
//                 method: "POST",
//                 headers: { Accept: "application/json;odata=verbose" },
//                 credentials: "same-origin"
//             }).then((response) => {
//                 return response.json();
//             });
//         } catch (error) {
//             console.log("getValues: " + error);
//         }
//     }

//     public getRequest(url: string): any {
//         try {
//             return fetch(url, {
//                 headers: { Accept: 'application/json;odata=verbose' },
//                 credentials: "same-origin"
//             }).then((response) => {
//                 if (response.status >= 200 && response.status < 400) {
//                     return response.json();
//                 }
//                 else {
//                     return response.json();
//                 }
//             }).catch(error =>
//                 console.error("getRequest: " + error));
//         } catch (error) {
//             console.log("getRequest: " + error);
//         }
//     }

//     //will send post request to add/update list data
//     public postRequest(url: string, postBody, xMethod, requestDigestValue): any {
//         try {
//             //GET FormDigestValue
//             return fetch(url, {
//                 headers: {
//                     Accept: 'application/json;odata=verbose',
//                     "Content-Type": 'application/json;odata=verbose',
//                     "X-RequestDigest": requestDigestValue,
//                     "X-Http-Method": xMethod,
//                     'IF-MATCH': '*'
//                 },
//                 method: 'POST',
//                 body: postBody,
//                 credentials: "same-origin"
//             }).then((response) => {
//                 if (response.status <= 204 && response.status >= 200) {
//                     //resolve(response);
//                     if (response.status == 204 || xMethod == 'DELETE') {
//                         return 'success';
//                     } else {
//                         return response.json();
//                     }
//                 }
//                 else {
//                     return response.json();
//                 }
//             }, (err) => {
//                 console.log(err);
//             }).catch(error =>
//                 console.error("postRequest: " + error));

//         } catch (error) {
//             console.log("postRequest: " + error);
//         }
//     }


//     public async uploadFiles(uploadFileObj, filename, context, ListDisplayName, siteurl, requestDigestValue) {
//         try {
//             if (uploadFileObj != '') {
//                 var file = uploadFileObj;
//                 if (file != undefined || file != null) {
//                     let spOpts: ISPHttpClientOptions = {
//                         headers: {
//                             "Accept": "application/json",
//                             "Content-Type": "application/json"
//                         },
//                         body: file,
//                         credentials: "same-origin"
//                     };
//                     var fileName = encodeURIComponent("" + filename + "");
//                     let url = `${siteurl}/_api/Web/Lists/getByTitle('${listTitle}')/RootFolder/Files/Add(url='${fileName}', overwrite=true)`;

//                     //POST call
//                     return context.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts).then(async (response: SPHttpClientResponse) => {
//                         return response.json();
//                     }).then(async (responseJSON: JSON) => {
//                         //updating columns values of this item
//                         return await this.updateLibraryColValue(responseJSON['Name'], ListDisplayName, siteurl, requestDigestValue);
//                     });
//                 }
//             } else {
//                 //updating columns values of this item
//                 return await this.updateLibraryColValue(fileName, ListDisplayName, siteurl, requestDigestValue);
//             }
//         }
//         catch (error) {
//             console.log('uploadFiles : ' + error);
//         }
//     }

//     //will update document library column value
//     public updateLibraryColValue(responseName, ListDisplayName, siteurl, requestDigestValue) {
//         try {
//             var itemURL = '';
//             var FileName = encodeURIComponent("" + responseName + "");
//             itemURL = `${siteurl}/_api/web/lists/getbytitle('${ListDisplayName}')/items?$select=FileRef,ID,LinkFilename&$filter=substringof('${FileName}',FileRef)`;
//             //GET request
//             return this.getRequest(itemURL).then(async (responseData) => {
//                 var resultData = responseData.d.results;
//                 if (resultData.length > 0) {
//                     var addDocColumnUrl = `${siteurl}/_api/web/lists/getbytitle('${ListDisplayName}')/items(${resultData[0].ID})`;
//                     var commomJSON = null;

//                     commomJSON = JSON.stringify({
//                         __metadata: { 'type': listEntityTypeName },
//                         Description0: this.state.postDescription

//                     });
//                     //POST request           
//                     return await this.postRequest(addDocColumnUrl, commomJSON, 'MERGE', requestDigestValue).then((data) => {
//                         return data;
//                     });
//                 }
//             });
//         } catch (error) {
//             console.log("postRequest: " + error);
//         }
//     }

//     public getCommunityPhotoDetails() {
//         promise = new Promise<any>((resolve, reject) => {
//             SPGet(
//                 siteURL +
//                 `/_api/web/lists/getByTitle('` +
//                 listTitle +
//                 `')/items?$select=*,Author/Id,Author/Title,FileRef,File/ServerRelativeUrl,FileLeafRef&$expand=File,Author&$orderby=Created desc&$top=1`
//             ).then(r => resolve(r));
//         });
//         promise
//             .then(r => {
//                 {
//                     this.setState({
//                         CommunityPhotosValues: r.value,
//                         isLoaded: true,
//                         currentItemID: r.value[0].Id
//                     });
//                     this.getCommunityLikesDetails();
//                 }
//             })
//             .catch(reject => {
//                 //console.log(reject);
//             });
//     }

//     public getCommunityLikesDetails() {
//         SPGet(
//             siteURL +
//             `/_api/web/lists/getByTitle('` +
//             Community_Like_List +
//             `')/items`
//         ).then(
//             async r =>
//                 await this.setState({
//                     CommunityLikeValues: r.value,
//                     isLoaded: true
//                 })
//         );
//     }

//     public stoploader() {
//         var loaderId = document.getElementById('dataLoader');
//         loaderId.style.display = 'none';
//     }

//     public startloader() {
//         var loaderId = document.getElementById('dataLoader');
//         loaderId.style.display = 'block';
//     }

//     public builddiv(SelectedUser) {
//         return (
//             <div className="section_innerwebpart">
//                 {/* <div className="dbb_mc_title_reply_row">
//           <div className="dbb_mctr_title">Replies</div>
//         </div> */}
//                 {this.buildRows(SelectedUser)}
//             </div>
//         );
//     }

//     public buildRows(SelectedUser) {
//         var reactHandler = this;
//         var getAndCheckcommentdata = [];
//         if (SelectedUser.ID != undefined && this.state.AllComments.length > 0) {
//             getAndCheckcommentdata = this.state.AllComments.filter(x =>
//                 x["CommunityUserId"] == SelectedUser.ID);
//         }

//         return (
//             <div className={getAndCheckcommentdata.length > 0 ? "discussion_board_box" : ["discussion_board_box", "no_comment"].join(' ')}>
//                 <div className="dbb_mc_input_row">
//                     <AutoComplateText SelectedUser={SelectedUser} siteURL={this.props.siteAbsoluteUrl} IsParent={true} selectedUserId={SelectedUser.ID} UserDetails={this.state.UserDetails} parentCallback={this.callbackFunction} />
//                     {this.state.IsReplyDone ? reactHandler.getDicussionCommentsAfterReply(SelectedUser.ID, SelectedUser) : ""}
//                 </div>

//                 <div className="dbb_reply_comment">
//                     <ul className="dbb_reply_comment_listing">
//                         {this.getreply(SelectedUser)}
//                     </ul>
//                 </div>
//             </div>
//         );
//     }

//     public getreply(SelectedUser) {
//         basesiteurl = window.location.origin;
//         basehuburl = basesiteurl + "/sites/contentTypeHub";
//         var replysymbol = basehuburl + "/Style%20Library/Community%20WebPart/images/reply.png";
//         var reactHandler = this;
//         return reactHandler.state.topicComments.map((item, index) => {
//             var replyDate = new Date(item['Created']);
//             var imgSrc = "";
//             return (
//                 <li>
//                     <div className="dbb_rcl_user_icon">
//                         <img src={this.props.siteAbsoluteUrl + "/_layouts/15/userphoto.aspx?size=S&accountname=" + item["Author"]["EMail"]} />
//                     </div>
//                     <div className="dbb_rcl_content_box">
//                         <div className="dbb_rcl_reply_text">
//                             <p className="replyCommentAuthorTitle">{item["Author"]['Title']}</p>
//                             {currentYear == replyDate.getUTCFullYear() ? <p className="replyCommentData">{monthNames[replyDate.getUTCMonth()] + " " + replyDate.getUTCDate()}</p> : <p className="replyCommentData">{monthNames[replyDate.getUTCMonth()] + " " + replyDate.getUTCDate() + ", " + replyDate.getUTCFullYear()}</p>}
//                             {/* Reply by {item["Author"]['Title']} on {monthNames[replyDate.getUTCMonth()] + " " + replyDate.getUTCDate() + ", " + replyDate.getUTCFullYear()} */}
//                         </div>

//                         <div>
//                             <div>
//                                 <p dangerouslySetInnerHTML={{ __html: item["Reply"] }} />
//                             </div>
//                         </div>

//                         <div className="dbb_rcl_reply_btn" onClick={(e) => this.toggleReplyDiv(e)}>
//                             <Icon className="ReplyIcon" iconName="Reply" />
//                             {/* <img src={replysymbol} /> */}
//                             <span className="ReplyText"> Reply</span>
//                         </div>

//                         <div className={["dbb_mc_input_row", "mp_hiddenDiv"].join(' ')}>
//                             <AutoComplateText SelectedUser={SelectedUser} siteURL={this.props.siteAbsoluteUrl} IsParent={false} selectedUserId={item["ID"]} UserDetails={this.state.UserDetails} parentCallback={this.callbackFunction} />
//                             {this.state.IsReplyDone ? reactHandler.getDicussionCommentsAfterReply(SelectedUser.ID, SelectedUser) : ""}
//                         </div>
//                         {this.getsubreply(reactHandler, index, monthNames)}
//                     </div>
//                 </li>
//             );
//         });
//     }

//     //Supportive Method
//     public toggleReplyDiv(e) {
//         // console.log('TSX: toggleReplyDiv');

//         var clickElement = e.target;
//         var parentDiv = clickElement.parentElement;
//         var divToShow = parentDiv.nextSibling;

//         if (divToShow.classList.contains("mp_hiddenDiv")) {
//             divToShow.className = divToShow.className.toString().substr(0, divToShow.className.toString().indexOf(' ' + "mp_hiddenDiv"));
//         }
//         else {
//             divToShow.className = divToShow.className.toString() + ' ' + "mp_hiddenDiv";
//         }
//     }

//     public getsubreply(reactHandler, index, monthNames) {
//         // console.log("TSX: getsubreply");

//         return reactHandler.state.topicComments[index].replies.map((replyItem, replyIndex) => {
//             var replyDate = new Date(replyItem.Created);
//             return (
//                 <ul className="dbb_reply_comment_sub_listing">
//                     <li>
//                         <div className="dbb_rcl_user_icon">
//                             <img src={this.props.siteAbsoluteUrl + "/_layouts/15/userphoto.aspx?size=S&accountname=" + replyItem["Author"]["EMail"]} />
//                         </div>
//                         <div className="dbb_rcl_content_box">
//                             <div className="dbb_rcl_reply_text">
//                                 <p className="replyCommentAuthorTitle">{replyItem.Author['Title']}</p>
//                                 {currentYear == replyDate.getUTCFullYear() ? <p className="replyCommentData">{monthNames[replyDate.getUTCMonth()] + " " + replyDate.getUTCDate()}</p> : <p className="replyCommentData">{monthNames[replyDate.getUTCMonth()] + " " + replyDate.getUTCDate() + ", " + replyDate.getUTCFullYear()}</p>}
//                                 {/* Reply by {replyItem.Author['Title']} on {monthNames[replyDate.getUTCMonth()] + " " + replyDate.getUTCDate() + ", " + replyDate.getUTCFullYear()} */}
//                             </div>
//                             <div className="dbb_rcl_content">
//                                 <p dangerouslySetInnerHTML={{ __html: replyItem.Reply }} />
//                             </div>
//                         </div>
//                     </li>
//                 </ul>
//             );
//         });
//     }

//     public async getLikeData(Id) {
//         countLike = countLike + 1;
//         if (countLike == 1) {
//             await fetch(siteURL + `/_api/Lists/getbytitle('` + Community_Like_List + `')/items?&select = *&$filter=CommunityPhotosIDId eq '` + Id + `' and Title eq '` + currentLoginUserInfo.userDisplayName.toLowerCase() + `'`, {
//                 credentials: 'same-origin', headers: { 'Accept': 'application/json;odata=verbose', 'odata-version': '' }
//             })
//                 .then((res) => res.json())
//                 .then(async (result) => {
//                     if (result.d.results.length <= 0) {
//                         await this.handleLikeClick(Id);
//                     }
//                     else {
//                         countLike = 0;
//                     }
//                 },
//                     (error) => {
//                         console.error('Error: _getKDCData', error);
//                     }
//                 );
//         }
//     }

//     public handleLikeClick(Id) {
//         promise = new Promise<any>((resolve, reject) => {
//             SPPost({
//                 url:
//                     siteURL + `/_api/Lists/getbytitle('` + Community_Like_List + `')/items`,
//                 payload: {
//                     Title: currentLoginUserInfo.userDisplayName,
//                     CommunityPhotosIDId: Id,
//                     IsLike: true
//                 }
//             }).then(r => resolve(r));
//         });
//         promise
//             .then(r => {
//                 {
//                     countLike = 0;
//                     this.getCommunityLikesDetails();
//                 }
//             })
//             .catch(reject => {
//                 //console.log(reject);
//             });
//     }

//     public async handleDsLikeClick(Id, likeID) {
//         siteURL = this.props.siteAbsoluteUrl;
//         var url = siteURL + `/_api/web/lists/getbytitle('` + Community_Like_List + `')/items(` + likeID + `)`;
//         let digest = await this.getRequestDigest(siteURL);
//         try {
//             fetch(url, {
//                 method: "POST",
//                 body: "{'IsDeleted': true }",
//                 headers: {
//                     credentials: "include",
//                     Accept: "application/json; odata=nometadata",
//                     "Content-Type": "application/json; odata=nometadata",
//                     "X-RequestDigest": digest.d.GetContextWebInformation.FormDigestValue,
//                     "IF-MATCH": "*",
//                     "X-HTTP-Method": "DELETE"
//                 }
//             })
//                 .then(
//                     (result) => {
//                         if (result.ok) {
//                             this.getCommunityLikesDetails();
//                         }
//                     }
//                 );
//         } catch (error) {
//             console.log("deletePost: " + error);
//         }

//     }

//     public getRequestDigest(siteURL) {
//         try {
//             var url = siteURL + "/_api/contextinfo";
//             return fetch(url, {
//                 method: "POST",
//                 headers: { Accept: "application/json;odata=verbose" }
//             })
//                 .then((res) => res.json())
//                 .then((response) => {
//                     return response;
//                 },
//                     (error) => {
//                         console.log('Error', error);
//                     });
//         } catch (error) {
//             console.log("getRequestDigest: " + error);
//         }
//     }

//     public async handleCommentClick(selectedItems) {
//         var Topiccomments: ITopicComments[] = [];
//         var AllComments = [];
//         if (this.state.AllComments.length > 0) {
//             this.state.AllComments.map((item, index) => {
//                 AllComments.push(item);
//                 if (item.Is_x0020_Parent == true && item.Thread_x0020_Parent == selectedItems.ID) {
//                     Topiccomments.push(item);
//                 }
//             });
//             Topiccomments.map((commentItem, commentIndex) => {
//                 var reply = [];
//                 AllComments.map((allCommentItem, allCommentIndex) => {
//                     if (allCommentItem.Is_x0020_Parent == false && allCommentItem.Thread_x0020_Parent == commentItem.ID) {
//                         reply.push(allCommentItem);
//                     }
//                 });
//                 Topiccomments[commentIndex]['replies'] = reply;
//             });
//             await this.setState({
//                 topicComments: Topiccomments,
//                 IsCommentClick: true,
//                 IsCommentUserID: selectedItems.ID,
//                 showDetailPopup: true,
//                 UserShownInModel: selectedItems,
//                 isLoaded: true
//             });

//         }
//         else {
//             await this.setState({
//                 topicComments: Topiccomments,
//                 IsCommentClick: true,
//                 IsCommentUserID: selectedItems.ID,
//                 showDetailPopup: true,
//                 UserShownInModel: selectedItems,
//                 isLoaded: true
//             });
//         }
//     }


//     public bindIcons(items) {
//         var cuItemID = this.state.currentItemID;
//         if (cuItemID != undefined && this.state.AllComments.length > 0) {
//             filterCommentData = this.state.AllComments.filter(x =>
//                 x["CommunityUserId"] == cuItemID);
//         }

//         var sortAndFilterData;
//         if (this.state.CommunityLikeValues.length > 0) {
//             sortAndFilterData = this.state.CommunityLikeValues.filter(x => x["CommunityPhotosIDId"] == cuItemID);
//             {
//                 cuItemID != undefined ? totalLike = sortAndFilterData : ""
//             }

//             if (sortAndFilterData.length > 0) {
//                 sortAndFilterData = sortAndFilterData.filter(x =>
//                     x["Title"].toLowerCase() == currentLoginUserInfo.userDisplayName.toLowerCase());
//                 if (sortAndFilterData.length > 0) {
//                     return (
//                         <div className="insta_likeComment">
//                             <div className="instaCommentBox">
//                                 <Icon onClick={() => this.handleCommentClick(items)} iconName="Comment" />
//                                 <span>{filterCommentData.length} comments</span>
//                             </div>

//                             <div className='instaLikeBox'>
//                                 <Icon className="insta_like_Icon_color" onClick={() => this.handleDsLikeClick(cuItemID, sortAndFilterData[0].ID)} iconName="LikeSolid" />
//                                 <span>{totalLike.length} Likes </span>
//                             </div>
//                         </div >);
//                 }
//                 else {
//                     return (
//                         <div className="insta_likeComment">
//                             <div className="instaCommentBox">
//                                 <Icon onClick={() => this.handleCommentClick(items)} iconName="Comment" />
//                                 <span>{filterCommentData.length} comments</span>
//                             </div>

//                             <div className='instaLikeBox'>
//                                 <Icon onClick={() => this.getLikeData(cuItemID)} iconName="Like" />
//                                 {/* <Icon onClick={() => this.handleLikeClick(cuItemID)} iconName="Like" /> */}
//                                 <span>{totalLike.length} Likes </span>
//                             </div>
//                         </div>
//                     );
//                 }
//             }
//             else {
//                 return (
//                     <div className="insta_likeComment">
//                         <div className="instaCommentBox">
//                             <Icon onClick={() => this.handleCommentClick(items)} iconName="Comment" />
//                             <span>{filterCommentData.length} comments</span>
//                         </div>

//                         <div className='instaLikeBox'>
//                             <Icon onClick={() => this.getLikeData(cuItemID)} iconName="Like" />
//                             {/* <Icon onClick={() => this.handleLikeClick(cuItemID)} iconName="Like" /> */}
//                             <span>{totalLike.length} Likes </span>
//                         </div>
//                     </div>
//                 );
//             }
//         }
//         else {
//             return (
//                 <div className="insta_likeComment">
//                     <div className="instaCommentBox">
//                         <Icon onClick={() => this.handleCommentClick(items)} iconName="Comment" />
//                         <span>{filterCommentData.length} comments</span>
//                     </div>

//                     <div className='instaLikeBox'>
//                         <Icon onClick={() => this.getLikeData(cuItemID)} iconName="Like" />
//                         {/* <Icon onClick={() => this.handleLikeClick(cuItemID)} iconName="Like" /> */}
//                         <span>{totalLike.length} Likes </span>
//                     </div>
//                 </div>
//             );
//         }

//     }

//     public closePopup(this) {
//         this.setState({
//             showDetailPopup: false,
//             isLoaded: true,
//             error: null,
//             UserShownInModel: "",
//             IsCommentClick: true,
//             IsReplyDone: false
//         });
//     }

//     public closeNewPostModal(this) {
//         this.setState({
//             showPopup: false,
//             postDescription: "",
//             postDscriptionErrorMessage: "",
//             fileToSend: "",
//             showFileErrorMessage: false,
//             fileErrorMessage: ""
//         });
//     }

//     callbackFunction = (childData) => {
//         this.setState({ IsReplyDone: childData })
//     }

//     public render(): React.ReactElement<ICommunityWebpartProps> {
//         const { error, isLoaded } = this.state;
//         let alldata = this.state.CommunityPhotosValues;
//         if (alldata.length > 0) {
//             return (
//                 <div>
//                     <div className="instagram_webpart">
//                         <div className="webpart_ttl">
//                             <p id="webpart_title_focus">{this.props.Title}
//                                 <PrimaryButton onClick={() => this.togglePopup()} text="Add Post" className="postImg_btn" />
//                             </p>
//                         </div>

//                         {alldata.map((item, index) => {
//                             var description = "";
//                             if (item.Description0 != null && item.Description0 != "" && item.Description0 != undefined) {
//                                 description = item.Description0.length > 165 ? item.Description0.substr(0, 80) + '...' : item.Description0;
//                             }
//                             return (
//                                 <div className="instaPost">
//                                     <div className="insta_img" onClick={() => this.handleCommentClick(item)}>
//                                         <img src={item.FileRef != "" ? item.FileRef + "?RenditionID=7" : window.location.origin + "/sites/contentTypeHub/Style%20Library/Community%20WebPart/images/noimage.png?RenditionID=7"} alt="Image Not Found" />
//                                     </div>
//                                     <div className="insta_footer">
//                                         {this.bindIcons(item)}
//                                     </div>
//                                     <div className="insta_description" onClick={() => this.handleCommentClick(item)}>{description}</div>
//                                     {item != "" ?
//                                         <div className="insta_setDate">
//                                             <span className="insta_AuthorName">{item.Author.Title}  </span>
//                                             <span>{monthNames[new Date(item.Created).getUTCMonth()] + " " + new Date(item.Created).getUTCDate() + ", " + new Date(item.Created).getUTCFullYear()}</span>
//                                         </div> : ""}
//                                     <div className="insta_footer">
//                                         {/* {this.bindIcons(item)} */}
//                                         {/* <PrimaryButton onClick={() => this.togglePopup()} text="POST IMAGE" className="postImg_btn" /> */}
//                                         <a href={this.props.siteAbsoluteUrl + "/SitePages/" + this.props.SeeAllUrl} target="_blank" data-interception="off">
//                                             View more photos <i className="fa fa-chevron-right" aria-hidden="true"></i>
//                                         </a>
//                                     </div>
//                                 </div>
//                             );
//                         })}

//                         {<Modal isOpen={this.state.showPopup} isBlocking={false}>
//                             <div className="modal_container kdc_comm_popup">
//                                 <div className="loader" id="dataLoader" style={{ display: 'none' }}>
//                                     <Spinner size={SpinnerSize.large} label='Loading...' />
//                                 </div>
//                                 <div className="modal_popup_box">
//                                     <div className="modal_title_box">
//                                         <div className="modal_title_close_btn" onClick={this.closeNewPostModal.bind(this)}>X</div>
//                                         <h3>Add New Post</h3>
//                                     </div>
//                                     <div className="modelContentWra">
//                                         {/* <AutoComplateText UserDetails={this.state.UserDetails} parentCallback={this.callbackFunction} />
//                     <p>{this.state.text}</p> */}
//                                         <div className="input_box_popup">
//                                             <label>Description</label>
//                                             <TextField placeholder="Description" multiline rows={4} onBlur={(e) => this.oneElementValueChange('postDescription', e, 'postDscriptionErrorMessage', 'Desc')} value={this.state.postDescription} />
//                                             <span className="errorMsg">{this.state.postDscriptionErrorMessage}</span>
//                                         </div>
//                                         <div className="input_box_popup">
//                                             <div className="uploadFileInput">
//                                                 <input type="file" id="FileUpload" value='' onChange={(e) => this.onFileUploadChange(e)} accept=".png,.jpg,.jpeg,.bmp,.gif" />
//                                                 <Icon iconName="Attach" />
//                                             </div>
//                                             <span className="selectedFile">{this.state.fileToSend}</span>
//                                             <span hidden={this.state.showFileErrorMessage} className="helperText">Only '.jpg, .jpeg, .png, .bmp, .gif' image file formats are supported.</span>
//                                             <span className="errorMsg">{this.state.fileErrorMessage}</span>
//                                         </div>
//                                     </div>
//                                     <div className="info_modal_btn-close kdcPostModalButtons">
//                                         {/* <PrimaryButton onClick={() => this.setState({ showPopup: false })} text="Cancel" /> */}
//                                         <DefaultButton onClick={this.closeNewPostModal.bind(this)} text="Cancel" />
//                                         <PrimaryButton onClick={this.validatePostsDetails.bind(this)} text="Save" />
//                                     </div>
//                                 </div>
//                             </div>
//                         </Modal>}

//                         <Modal isOpen={this.state.showDetailPopup} isBlocking={false}>
//                             <div className="modal_container kdc_comm_popup">
//                                 <div className="modal_popup_box">
//                                     <div className="modal_title_box">
//                                         <h3 id="model_Title_autofocus">Photo Details</h3>
//                                         <div style={{ visibility: 'hidden' }}>
//                                             <TextField autoFocus={false} />
//                                         </div>
//                                         <div className="modal_title_close_btn" onClick={this.closePopup.bind(this)} >X</div>
//                                     </div>
//                                     <div className="modelContentWra">
//                                         <div className="pa_pContentMain">
//                                             <div className="paModalImage">
//                                                 <img src={this.state.UserShownInModel.FileRef != "" ? this.state.UserShownInModel.FileRef : window.location.origin + "/sites/contentTypeHub/Style%20Library/Community%20WebPart/images/noimage.png"} />
//                                             </div>
//                                             <div className="pa_contentIcon">
//                                                 {this.bindIcons(this.state.showDetailPopup)}
//                                             </div>
//                                             <div className="pa_contentDescri">
//                                                 {this.state.UserShownInModel.Description0}
//                                             </div>
//                                             {this.state.UserShownInModel != "" ?
//                                                 <div className="model_setDate">
//                                                     <span className="model_AuthorName">{this.state.UserShownInModel.Author.Title}  </span>
//                                                     <span>{monthNames[new Date(this.state.UserShownInModel.Created).getUTCMonth()] + " " + new Date(this.state.UserShownInModel.Created).getUTCDate() + ", " + new Date(this.state.UserShownInModel.Created).getUTCFullYear()}</span>
//                                                 </div> : ""}
//                                             {this.state.IsCommentClick && this.state.IsCommentUserID == this.state.UserShownInModel.ID ? this.builddiv(this.state.UserShownInModel) : ""}
//                                         </div>
//                                     </div>
//                                     <div className="info_modal_btn-close kdcPostModalButtons">
//                                         {/* <DefaultButton className="vp_close_btn" onClick={this.closePopup.bind(this)} text="Close" /> */}
//                                     </div>
//                                 </div>
//                             </div>
//                         </Modal>

//                     </div>
//                     {/* {this.state.IsCommentClick ? this.builddiv() : ""} */}
//                 </div>
//             );
//         } else if (!isLoaded) {
//             return <Spinner size={SpinnerSize.large} label="Loading..." />;
//         } else if (error) {
//             return <div>Error: {error.message}</div>;
//         } else {
//             return (
//                 <div>
//                     <div className="instagram_webpart">
//                         {/* <div className="titleMain">
//               <h3>{this.props.Title}</h3>
//               <a href={this.props.siteAbsoluteUrl + "/SitePages/" + this.props.SeeAllUrl} target="_blank" data-interception="off">
//                 View All Photos <i className="fa fa-chevron-right" aria-hidden="true"></i>
//               </a>
//             </div> */}
//                         <div className="webpart_ttl">
//                             <p>{this.props.Title}
//                                 <PrimaryButton onClick={() => this.togglePopup()} text="Add Post" className="postImg_btn" />
//                             </p>
//                         </div>
//                         <br></br>
//                         <br></br>
//                         <div>Post Not Found. Please Enter New Post.</div>
//                         <br></br>
//                         <div className="insta_footer">
//                             <a href={this.props.siteAbsoluteUrl + "/SitePages/" + this.props.SeeAllUrl} target="_blank" data-interception="off">
//                                 View more photos <i className="fa fa-chevron-right" aria-hidden="true"></i>
//                             </a>
//                         </div>
//                         {<Modal isOpen={this.state.showPopup} isBlocking={false}>
//                             <div className="modal_container kdc_comm_popup">
//                                 <div className="modal_popup_box">
//                                     <div className="modal_title_box">
//                                         <div className="modal_title_close_btn" onClick={this.closeNewPostModal.bind(this)}>X</div>
//                                         <h3>Add New Post</h3>
//                                     </div>
//                                     <div className="modelContentWra">
//                                         {/* <AutoComplateText UserDetails={this.state.UserDetails} /> */}
//                                         <div className="input_box_popup">
//                                             <label>Description</label>
//                                             <TextField placeholder="Description" multiline rows={4} onBlur={(e) => this.oneElementValueChange('postDescription', e, 'postDscriptionErrorMessage', 'Desc')} value={this.state.postDescription} />
//                                             <span className="errorMsg">{this.state.postDscriptionErrorMessage}</span>
//                                         </div>
//                                         <div className="input_box_popup">
//                                             <div className="uploadFileInput">
//                                                 <input type="file" id="FileUpload" value='' onChange={(e) => this.onFileUploadChange(e)} accept=".png,.jpg,.jpeg,.bmp" />
//                                             </div>
//                                             <span className="selectedFile">{this.state.fileToSend}</span>
//                                             <span hidden={this.state.showFileErrorMessage} className="errorMsg">Only '.jpg, .jpeg, .png, .bmp' image file formats are supported.</span>
//                                             <span className="errorMsg">{this.state.fileErrorMessage}</span>
//                                         </div>
//                                     </div>
//                                     <div className="info_modal_btn-close kdcPostModalButtons">
//                                         <PrimaryButton onClick={this.closeNewPostModal.bind(this)} text="Cancel" />
//                                         <PrimaryButton onClick={this.validatePostsDetails.bind(this)} text="Save" />
//                                     </div>
//                                 </div>
//                             </div>
//                         </Modal>}
//                     </div>
//                 </div>
//             );
//         }
//     }
// }

