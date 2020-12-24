import { SPPermission } from "@microsoft/sp-page-context";

export const getValues = (webURL) => {
    try {
        var url = webURL + "/_api/contextinfo";
        return fetch(url, {
            method: "POST",
            credentials: 'same-origin',
            headers: { Accept: "application/json;odata=verbose" }
        })
            .then((response) => {
                var datum = response.json();
                console.log(datum);
                return datum;
            });
    } catch (error) {
        console.log("getValues: " + error);
    }
};


export const getListItems = async (url, currentLoggedInUser) => {
    let response = await fetch(url, {
        method: 'GET',
        mode: 'cors',
        credentials: 'same-origin',
        headers: new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache',
            'pragma': 'no-cache',
        }),
    });

    const responseString = await response.json();
    var resultJsonPasrse = await JSON.stringify(responseString);
    var data = await JSON.parse(resultJsonPasrse);
    console.log(data);
    if (data.value.length > 0) {
        return data.value[0];
    }
    else {
        return false;
    }
};





export const getListItem = async (url) => {
    let response = await fetch(url, {
        method: 'GET',
        mode: 'cors',
        credentials: 'same-origin',
        headers: new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache',
            'pragma': 'no-cache',
        }),
    });

    const responseString = await response.json();
    var resultJsonPasrse = await JSON.stringify(responseString);
    var data = await JSON.parse(resultJsonPasrse);
    console.log(data);
    return data;
};

export const deleteAttachments = async (webURL, itemID, requestDigestValue, FilesToDelete, delCnt) => {
    if (delCnt < FilesToDelete.length) {
        let FileURL = webURL + "/_api/web/lists/GetByTitle('StudentInformation')/GetItemById(" + itemID + ")/AttachmentFiles/getByFileName('" + FilesToDelete[delCnt] + "') ";
        let response = await fetch(FileURL, {
            method: 'POST',
            mode: 'cors',
            credentials: 'same-origin',
            headers: new Headers({
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Access-Control-Allow-Origin': '*',
                'Cache-Control': 'no-cache',
                'pragma': 'no-cache',
                'X-HTTP-Method': 'DELETE',
                "X-RequestDigest": requestDigestValue,
                'If-Match': '*',
            }),
        });

        console.log(response.ok);
        if (response.ok) {
            delCnt++;
            deleteAttachments(webURL, itemID, requestDigestValue, FilesToDelete, delCnt);
        }
        else {
            delCnt++;
            deleteAttachments(webURL, itemID, requestDigestValue, FilesToDelete, delCnt);
        }
        return response.ok;
    }
};

export const callUploadFile = async (webURL, Files, requestDigestValue, itemId, cnt, FilesTageName) => {
    if (cnt < Files.length) {
        await UploadFile(webURL, Files, requestDigestValue, itemId, cnt, FilesTageName);
    }
};

export const updateData = async (webURL, url, body, requestDigestValue, FilesToDelete, Files, itemId, FilesTageName) => {
    let response = await fetch(url, {
        method: 'POST',
        mode: 'cors',
        credentials: 'same-origin',
        headers: new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache',
            'pragma': 'no-cache',
            'X-HTTP-Method': 'MERGE',
            "X-RequestDigest": requestDigestValue,
            'If-Match': '*',
        }),
        body: body
    });

    console.log(response.ok);
    if (response.ok) {
        await deleteAttachments(webURL, itemId, requestDigestValue, FilesToDelete, 0);
        await callUploadFile(webURL, Files, requestDigestValue, itemId, 0, FilesTageName);
    }
    return response.ok;
};



export const UploadFile = async (webURL, Files, requestDigestValue, itemID, cnt, FilesTageName) => {
    var cntVal = cnt;
    // var fileName = Files[cntVal].name;
    var fileName = FilesTageName[cntVal] + Files[cntVal].name;
    var url = webURL + "/_api/web/lists/GetByTitle('StudentInformation')/items(" + itemID + ")/AttachmentFiles/add(FileName='" + fileName + "')";
    fetch(url, {
        credentials: 'same-origin',
        method: "POST",
        headers: new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache',
            'pragma': 'no-cache',
            "X-RequestDigest": requestDigestValue,
        }),
        body: Files[cntVal]
    })
        .then((attchResult) => {
            cntVal++;
            callUploadFile(webURL, Files, requestDigestValue, itemID, cntVal, FilesTageName);
            console.log(attchResult);
        });
};

export const postData = async (webURL, url, body, requestDigestValue, Files, FilesTageName) => {
    fetch(url, {
        method: 'POST',
        mode: 'cors',
        credentials: 'same-origin',
        headers: new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache',
            'pragma': 'no-cache',
            "X-RequestDigest": requestDigestValue,

        }),
        body: body,
    }).then(async (response) => {
        var responseResult = await response.json();
        console.log(responseResult);
        callUploadFile(webURL, Files, requestDigestValue, responseResult.ID, 0, FilesTageName);
        return responseResult;
    }).catch((error) => {
        console.log("getValues: " + error);
    });
};

// export const _getYears = async (webURL) => {
//     let resultarr = [];
//     //var years1=[];
//     await fetch(webURL + "/_api/web/lists/getbytitle('Employee Financial Year')/Items", {
//         method: 'GET',
//         mode: 'cors',
//         credentials:'same-origin',
//         headers: new Headers({
//             'Content-Type': 'application/json',
//             'Accept': 'application/json',
//             'Access-Control-Allow-Origin': '*',
//             'Cache-Control': 'no-cache',
//             'pragma': 'no-cache',
//         }),
//     }).then((response) => response.json())
//         .then((data) => {
//              for (var i = 0; i < data.value.length; i++) {
//                  resultarr.push({
//                     ID: data.value[i].ID.toString(),
//                     years:data.value[i].Title,

//                     //ID:data.value[i].ID,

//                     //left side pannumber varialble must be same as field name of columns constant in tsx file

//                 });
//             }
//         });

//     return resultarr;
// };

export const _getPermission = async (webURL) => {
    var permission = "";
    //var years1=[];
    await fetch(webURL + "/_api/web/currentUser/issiteadmin", {
        method: 'GET',
        mode: 'cors',
        credentials: 'same-origin',
        headers: new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache',
            'pragma': 'no-cache',
        }),
    }).then((response) => response.json())
        .then((data) => {

            permission = data.value;

        });

    // return resultarr;
    return permission;
};

export const _getEffectiveBasePermission = async (webURL) => {
    var permission = [];

    await fetch(webURL + "/_api/web/lists/getbytitle('StudentInformation')/EffectiveBasePermissions", {
        method: 'GET',
        mode: 'cors',
        credentials: 'same-origin',
        headers: new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache',
            'pragma': 'no-cache',
        }),
    }).then((response) => response.json())
        .then((data) => {

            permission.push(new SPPermission(data));


        });

    return permission[0];
};

export const getAllData = async (webURL) => {
    let resultarr = [];
    await fetch(webURL + "/_api/web/lists/getbytitle('StudentInformation')/Items?$select=*,AttachmentFiles,AttachmentFiles/ServerRelativeUrl&$expand=AttachmentFiles", {
        method: 'GET',
        mode: 'cors',
        credentials: 'same-origin',
        headers: new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache',
            'pragma': 'no-cache',
        }),
    }).then((response) => response.json())
        .then(async (data) => {
            for (var i = 0; i < data.value.length; i++) {
                resultarr.push({
                    ID: data.value[i].ID,
                    Title: data.value[i].Title,
                    DateOfJoining: data.value[i].DateOfJoining,
                    //left side pannumber varialble must be same as field name of columns constant in tsx file
                    Created: data.value[i].Created,
                    PANnumber: data.value[i].PANnumber,
                    PaymentOfHouseRent: data.value[i].TotalA,
                    Financial_x0020_Year: data.value[i].Financial_x0020_Year,
                    TotalA: data.value[i].TotalA,
                    MedicalInsurancePremiumSelf: data.value[i].MedicalInsurancePremiumSelf,
                    MedicalInsurancePremiumParents: data.value[i].MedicalInsurancePremiumParents,
                    TotalB: data.value[i].TotalB,
                    InterestPaymentOnHousingLoan: data.value[i].InterestPaymentOnHousingLoan,
                    TotalC: data.value[i].TotalC,
                    UnitedLinkInsurancePlan: data.value[i].UnitedLinkInsurancePlan,
                    PublicProvidentFund: data.value[i].PublicProvidentFund,
                    LifeInsurancePremium: data.value[i].LifeInsurancePremium,
                    NationalSavingsCertificate: data.value[i].NationalSavingsCertificate,
                    UnitesOfApprovedMutualFund: data.value[i].UnitesOfApprovedMutualFund,
                    TermDeposit: data.value[i].TermDeposit,
                    HousingLoan: data.value[i].HousingLoan,
                    ChildrenSchoolFees: data.value[i].ChildrenSchoolFees,
                    NotifiedPensionScheme: data.value[i].NotifiedPensionScheme,
                    TotalD: data.value[i].TotalD,
                    HouseRent_Jan: data.value[i].HouseRent_Jan,
                    HouseRent_Jan_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Jan_Doc,
                    HouseRent_Feb_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Feb_Doc,
                    HouseRent_March_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_March_Doc,
                    HouseRent_April_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_April_Doc,
                    HouseRent_May_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_May_Doc,
                    HouseRent_June_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_June_Doc,
                    HouseRent_July_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_July_Doc,
                    HouseRent_Aug_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Aug_Doc,
                    HouseRent_Sept_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Sept_Doc,
                    HouseRent_Oct_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Oct_Doc,
                    HouseRent_Nov_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Nov_Doc,
                    HouseRent_Dec_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Dec_Doc,
                    MedicalInsurancePremiumSelf_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].MedicalInsurancePremiumSelf_Doc,
                    MedicalInsurancePremiumParents_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].MedicalInsurancePremiumParents_D,
                    InterestPaymentOnHousingLoan_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].InterestPaymentOnHousingLoan_Doc,
                    UnitedLinkInsurancePlan_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].UnitedLinkInsurancePlan_Doc,
                    PublicProvidentFund_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].PublicProvidentFund_Doc,
                    LifeInsurancePremium_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].LifeInsurancePremium_Doc,
                    NationalSavingsCertificate_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].NationalSavingsCertificate_Doc,
                    UnitesOfApprovedMutualFund_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].UnitesOfApprovedMutualFund_Doc,
                    TermDeposit_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].TermDeposit_Doc,
                    HousingLoan_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HousingLoan_Doc,
                    ChildrenSchoolFees_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].ChildrenSchoolFees_Doc,
                    NotifiedPensionScheme_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].NotifiedPensionScheme_Doc,
                    RentAgreementAttachment_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].RentAgreementAttachment,
                    HouseRent_Feb: data.value[i].HouseRent_Feb,
                    HouseRent_March: data.value[i].HouseRent_March,
                    HouseRent_April: data.value[i].HouseRent_April,
                    HouseRent_May: data.value[i].HouseRent_May,
                    HouseRent_June: data.value[i].HouseRent_June,
                    HouseRent_July: data.value[i].HouseRent_July,
                    HouseRent_Aug: data.value[i].HouseRent_Aug,
                    HouseRent_Sept: data.value[i].HouseRent_Sept,
                    HouseRent_Oct: data.value[i].HouseRent_Oct,
                    HouseRent_Nov: data.value[i].HouseRent_Nov,
                    HouseRent_Dec: data.value[i].HouseRent_Dec,
                    HouseRent_July_Doc: data.value[i].HouseRent_July_Doc,
                    HouseRent_Jan_Doc: data.value[i].HouseRent_Jan_Doc,
                    HouseRent_Feb_Doc: data.value[i].HouseRent_Feb_Doc,
                    HouseRent_March_Doc: data.value[i].HouseRent_March_Doc,
                    HouseRent_April_Doc: data.value[i].HouseRent_April_Doc,
                    HouseRent_May_Doc: data.value[i].HouseRent_May_Doc,
                    HouseRent_June_Doc: data.value[i].HouseRent_June_Doc,
                    HouseRent_Aug_Doc: data.value[i].HouseRent_Aug_Doc,
                    HouseRent_Sept_Doc: data.value[i].HouseRent_Sept_Doc,
                    HouseRent_Oct_Doc: data.value[i].HouseRent_Oct_Doc,
                    HouseRent_Nov_Doc: data.value[i].HouseRent_Nov_Doc,
                    HouseRent_Dec_Doc: data.value[i].HouseRent_Dec_Doc,
                    MedicalInsurancePremiumSelf_Doc: data.value[i].MedicalInsurancePremiumSelf_Doc,
                    MedicalInsurancePremiumParents_Doc: data.value[i].MedicalInsurancePremiumParents_D,
                    InterestPaymentOnHousingLoan_Doc: data.value[i].InterestPaymentOnHousingLoan_Doc,
                    UnitedLinkInsurancePlan_Doc: data.value[i].UnitedLinkInsurancePlan_Doc,
                    PublicProvidentFund_Doc: data.value[i].PublicProvidentFund_Doc,
                    LifeInsurancePremium_Doc: data.value[i].LifeInsurancePremium_Doc,
                    NationalSavingsCertificate_Doc: data.value[i].NationalSavingsCertificate_Doc,
                    UnitesOfApprovedMutualFund_Doc: data.value[i].UnitesOfApprovedMutualFund_Doc,
                    TermDeposit_Doc: data.value[i].TermDeposit_Doc,
                    HousingLoan_Doc: data.value[i].HousingLoan_Doc,
                    ChildrenSchoolFees_Doc: data.value[i].ChildrenSchoolFees_Doc,
                    NotifiedPensionScheme_Doc: data.value[i].NotifiedPensionScheme_Doc,
                    LandlordPANnumber: data.value[i].LandlordPANnumber,
                    RentAgreementAttachment: data.value[i].RentAgreementAttachment
                });
            }
        });
    return resultarr;
};


export const getAPI = async (webURL, Year) => {
    let resultarr = [];
    await fetch(webURL + "/_api/web/lists/GetByTitle('StudentInformation')/items?$select=*,Author/Title&$expand=Author&$filter=Financial_x0020_Year eq '" + Year + "'", {
        method: 'GET',
        mode: 'cors',
        credentials: 'same-origin',
        headers: new Headers({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Cache-Control': 'no-cache',
            'pragma': 'no-cache',
        }),
    }).then((response) => response.json())
        .then(async (data) => {
            for (var i = 0; i < data.value.length; i++) {
                resultarr.push({
                    ID: data.value[i].ID,
                    Title: data.value[i].Title,
                    DateOfJoining: data.value[i].DateOfJoining,
                    //left side pannumber varialble must be same as field name of columns constant in tsx file
                    Created: data.value[i].Created,
                    PANnumber: data.value[i].PANnumber,
                    PaymentOfHouseRent: data.value[i].TotalA,
                    Financial_x0020_Year: data.value[i].Financial_x0020_Year,
                    TotalA: data.value[i].TotalA,
                    MedicalInsurancePremiumSelf: data.value[i].MedicalInsurancePremiumSelf,
                    MedicalInsurancePremiumParents: data.value[i].MedicalInsurancePremiumParents,
                    TotalB: data.value[i].TotalB,
                    InterestPaymentOnHousingLoan: data.value[i].InterestPaymentOnHousingLoan,
                    TotalC: data.value[i].TotalC,
                    UnitedLinkInsurancePlan: data.value[i].UnitedLinkInsurancePlan,
                    PublicProvidentFund: data.value[i].PublicProvidentFund,
                    LifeInsurancePremium: data.value[i].LifeInsurancePremium,
                    NationalSavingsCertificate: data.value[i].NationalSavingsCertificate,
                    UnitesOfApprovedMutualFund: data.value[i].UnitesOfApprovedMutualFund,
                    TermDeposit: data.value[i].TermDeposit,
                    HousingLoan: data.value[i].HousingLoan,
                    ChildrenSchoolFees: data.value[i].ChildrenSchoolFees,
                    NotifiedPensionScheme: data.value[i].NotifiedPensionScheme,
                    TotalD: data.value[i].TotalD,
                    HouseRent_Jan: data.value[i].HouseRent_Jan,
                    HouseRent_Jan_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Jan_Doc,
                    HouseRent_Feb_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Feb_Doc,
                    HouseRent_March_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_March_Doc,
                    HouseRent_April_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_April_Doc,
                    HouseRent_May_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_May_Doc,
                    HouseRent_June_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_June_Doc,
                    HouseRent_July_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_July_Doc,
                    HouseRent_Aug_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Aug_Doc,
                    HouseRent_Sept_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Sept_Doc,
                    HouseRent_Oct_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Oct_Doc,
                    HouseRent_Nov_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Nov_Doc,
                    HouseRent_Dec_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HouseRent_Dec_Doc,
                    MedicalInsurancePremiumSelf_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].MedicalInsurancePremiumSelf_Doc,
                    MedicalInsurancePremiumParents_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].MedicalInsurancePremiumParents_D,
                    InterestPaymentOnHousingLoan_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].InterestPaymentOnHousingLoan_Doc,
                    UnitedLinkInsurancePlan_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].UnitedLinkInsurancePlan_Doc,
                    PublicProvidentFund_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].PublicProvidentFund_Doc,
                    LifeInsurancePremium_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].LifeInsurancePremium_Doc,
                    NationalSavingsCertificate_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].NationalSavingsCertificate_Doc,
                    UnitesOfApprovedMutualFund_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].UnitesOfApprovedMutualFund_Doc,
                    TermDeposit_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].TermDeposit_Doc,
                    HousingLoan_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].HousingLoan_Doc,
                    ChildrenSchoolFees_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].ChildrenSchoolFees_Doc,
                    NotifiedPensionScheme_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].NotifiedPensionScheme_Doc,
                    RentAgreementAttachment_Doc_URL: webURL + '/Lists/Employee%20Investment%20Declaration/Attachments/' + data.value[i].ID + '/' + data.value[i].RentAgreementAttachment,
                    HouseRent_Feb: data.value[i].HouseRent_Feb,
                    HouseRent_March: data.value[i].HouseRent_March,
                    HouseRent_April: data.value[i].HouseRent_April,
                    HouseRent_May: data.value[i].HouseRent_May,
                    HouseRent_June: data.value[i].HouseRent_June,
                    HouseRent_July: data.value[i].HouseRent_July,
                    HouseRent_Aug: data.value[i].HouseRent_Aug,
                    HouseRent_Sept: data.value[i].HouseRent_Sept,
                    HouseRent_Oct: data.value[i].HouseRent_Oct,
                    HouseRent_Nov: data.value[i].HouseRent_Nov,
                    HouseRent_Dec: data.value[i].HouseRent_Dec,
                    HouseRent_July_Doc: data.value[i].HouseRent_July_Doc,
                    HouseRent_Jan_Doc: data.value[i].HouseRent_Jan_Doc,
                    HouseRent_Feb_Doc: data.value[i].HouseRent_Feb_Doc,
                    HouseRent_March_Doc: data.value[i].HouseRent_March_Doc,
                    HouseRent_April_Doc: data.value[i].HouseRent_April_Doc,
                    HouseRent_May_Doc: data.value[i].HouseRent_May_Doc,
                    HouseRent_June_Doc: data.value[i].HouseRent_June_Doc,
                    HouseRent_Aug_Doc: data.value[i].HouseRent_Aug_Doc,
                    HouseRent_Sept_Doc: data.value[i].HouseRent_Sept_Doc,
                    HouseRent_Oct_Doc: data.value[i].HouseRent_Oct_Doc,
                    HouseRent_Nov_Doc: data.value[i].HouseRent_Nov_Doc,
                    HouseRent_Dec_Doc: data.value[i].HouseRent_Dec_Doc,
                    MedicalInsurancePremiumSelf_Doc: data.value[i].MedicalInsurancePremiumSelf_Doc,
                    MedicalInsurancePremiumParents_Doc: data.value[i].MedicalInsurancePremiumParents_D,
                    InterestPaymentOnHousingLoan_Doc: data.value[i].InterestPaymentOnHousingLoan_Doc,
                    UnitedLinkInsurancePlan_Doc: data.value[i].UnitedLinkInsurancePlan_Doc,
                    PublicProvidentFund_Doc: data.value[i].PublicProvidentFund_Doc,
                    LifeInsurancePremium_Doc: data.value[i].LifeInsurancePremium_Doc,
                    NationalSavingsCertificate_Doc: data.value[i].NationalSavingsCertificate_Doc,
                    UnitesOfApprovedMutualFund_Doc: data.value[i].UnitesOfApprovedMutualFund_Doc,
                    TermDeposit_Doc: data.value[i].TermDeposit_Doc,
                    HousingLoan_Doc: data.value[i].HousingLoan_Doc,
                    ChildrenSchoolFees_Doc: data.value[i].ChildrenSchoolFees_Doc,
                    NotifiedPensionScheme_Doc: data.value[i].NotifiedPensionScheme_Doc,
                    LandlordPANnumber: data.value[i].LandlordPANnumber,
                    RentAgreementAttachment: data.value[i].RentAgreementAttachment
                });
            }
        });
    return resultarr;
};

