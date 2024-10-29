import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { ConfidentialClientApplication } from '@azure/msal-node';

const ACCOUNT_ID = ""; //Replace with account Id
const PAYHAWK_API_KEY = ""; //Replace with API Key
const PAYHAWK_API_BASE_URL = `https://api.payhawk.com/api/v3/accounts/${ACCOUNT_ID}`;

const driveId = ''; // Replace with Sharepoint Drive ID
const templateFileId = ''; // Replace with Template File ID
const realFilesFolderId = ''; // Replace with folder ID

const msalConfig = {
    auth: {
        clientId: '', // Replace with your client ID
        authority: 'https://login.microsoftonline.com/TENANT', // Replace with your tenant ID
        clientSecret: '' // Replace with your client secret
    }
};

const PHLDR_TRIP_REASON = 'trip_reason';                    // Replace with Custom Field Value 
const PHLDR_TRANSPORT_TYPE = 'transport_type';              // Replace with Custom Field Value
// add more if needed 

const GRAPH_API_BASE_URL = 'https://graph.microsoft.com/v1.0';
const PHLDR_EXPENSE_ID = 'expense_id';                      //ID of the expense with 5 trailing zeros
const PHLDR_FROM_DATE = 'from_date';                        //First day of the business trip in format dd.mm.yyyy
const PHLDR_TO_DATE = 'to_date';                            //Last day of the business trip in format dd.mm.yyyy
const PHLDR_DESTINATION = 'destination';                    //The destination of the trip
const PHLDR_EXPENSE_DATE = 'expense_created_date';          //Date of the expense in format 22 Януари 2024
const PHLDR_EMPLOYEE_NAME = 'employee_name';                //Name of the employee that created the expense (in cyrillic as taken from the employee reimbursement details)
const PHLDR_EMPLOYEE_TEAM = 'employee_team';                //The team of the employee
const PHLDR_EMPLOYEE_PARENT_TEAM = 'employee_parent_team';  //The parent team of the team of the employee
const PHLDR_APPROVER_NAME = 'approver_name';                //Name of the employee that approves the expense (in cyrillic as taken from the employee reimbursement details)
const PHLDR_TRIP_TOTAL_AMOUNT = 'trip_total_amount';        //The total amount and currency for the trip


export async function payhawkWebhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log(`Http function processed request for url "${request.url}"`);

    let bodyJson = await request.json();
    let bodyString = JSON.stringify(bodyJson);

    let expenseId = bodyJson["payload"]["expenseId"] ?? '';
    if (!expenseId) {
        return { body: `${`No Expense id found in: ${bodyString}`}!` };
    }

    try {
        //Get Expense
        const expense = await getExpense(expenseId);
        const expenseData = await getExpenseData(expense);

        let azureAccessToken = await getAccessToken();
        const newFileName = `Per_Diem_For_Expense_${expenseId}.docx`;

        // Generate a new file
        const newFileResponse = await copyAndRenameFile(
            driveId,
            templateFileId,
            realFilesFolderId,
            newFileName,
            azureAccessToken,
        );

        let newFileId = null;
        if (newFileResponse.headers.get('location')) {
            newFileId = extractIdFromUrl(newFileResponse.headers.get('location'))
        }
        if (newFileId == null) {
            throw "Unable to retrieve ID of the file"
        }

        // replace placeholders
        const placeholders = {
            expense_id: expense.id,
            expense_created_date: expenseData[PHLDR_EXPENSE_DATE],
            employee_name: `${expense.createdBy.firstName} ${expense.createdBy.lastName}`,
            employee_parent_team: expenseData[PHLDR_EMPLOYEE_TEAM],
            destination: expenseData[PHLDR_DESTINATION],
            trip_reason: expenseData[PHLDR_TRIP_REASON],
            from_date: expenseData[PHLDR_FROM_DATE],
            to_date: expenseData[PHLDR_TO_DATE],
            approver_name: expenseData.PHLDR_APPROVER_NAME,
        };

        await replacePlaceholdersInFile(
            driveId,
            realFilesFolderId,
            newFileId,
            placeholders,
            azureAccessToken
        );

        // download as pdf
        let buffer = await downloadFileAsPDF(newFileId, driveId, azureAccessToken);

        // upload to Payhawk
        let result = await uploadDocumentToPayhawk(expense.id, buffer, newFileName);
    }
    catch (e) {
        console.log("Exepction Occured :");
        console.log(e);
        return { status: 500, body: `${`Exception`} ${e}!` };
    }

    return { body: `${`Successfully executed`}!` };
};

app.http('payhawkWebhook', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: payhawkWebhook
});

async function getExpense(expenseId: string): Promise<any> {
    try {
        const expenseResponse = await makePayhawkHttpRequest('GET', `expenses/${expenseId}`, null, null);
        return expenseResponse;
    } catch (error) {
        console.log(error);
    }
}

async function getExpenseData(expense: any): Promise<any> {
    const expenseId = expense.id;
    const result: any = {};
    try {
        const expenseWorkflowResponse = await makePayhawkHttpRequest('GET', `expenses/${expenseId}/workflow`, null, null);
        const expenseWorkflow = expenseWorkflowResponse;

        //Expense ID
        var myformat = new Intl.NumberFormat('en-US', {
            minimumIntegerDigits: 7,
            minimumFractionDigits: 0,
            useGrouping: false
        });
        result[PHLDR_EXPENSE_ID] = myformat.format(expense.id);

        //First and last date of the business trip
        const stops = expense.perDiem?.stops;
        const firstStop = stops[0];
        const lastStop = stops[stops.length - 1];
        result[PHLDR_FROM_DATE] = formatShortDate(new Date(Date.parse(firstStop.date)));
        result[PHLDR_TO_DATE] = formatShortDate(new Date(Date.parse(lastStop.date)));

        //Expense created date
        const createdAt: Date = new Date(Date.parse(expense.createdAt));
        result[PHLDR_EXPENSE_DATE] = formatLongDate(createdAt);

        //Employee name
        result[PHLDR_EMPLOYEE_NAME] = await getEmployeeName(expense.createdBy);

        //Trip total amount
        var totalAmountFormat = new Intl.NumberFormat('en-US', {
            minimumIntegerDigits: 1,
            minimumFractionDigits: 2,
            useGrouping: false
        });
        result[PHLDR_TRIP_TOTAL_AMOUNT] = totalAmountFormat.format(expense.reconciliation.totalAmount);


        //Approver name
        if (expenseWorkflow.approvedBy) {
            result[PHLDR_APPROVER_NAME] = await getEmployeeName(expenseWorkflow.approvedBy);
        } else {
            result[PHLDR_APPROVER_NAME] = 'No Approver';
        }

        //Destination
        let destinationString = '';
        if (stops.length === 2) {
            destinationString = `${stops[1].address}`;
        } else {
            for (let i = 1; i < stops.length; i++) {
                destinationString += `${i}. ${stops[i].address}\n`;
            }
        }
        result[PHLDR_DESTINATION] = destinationString;

        const customFields: [any] = expense.reconciliation.customFields;
        customFields.forEach(customField => {
            switch (customField.id) {
                case 'teams': //Employee team and parent team
                    result[PHLDR_EMPLOYEE_TEAM] = customField.selectedValues[0].label;
                    result[PHLDR_EMPLOYEE_PARENT_TEAM] = customField.selectedValues.length === 2 ? customField.selectedValues[1].label : '';
                    break;
                case PHLDR_TRIP_REASON: //Trip reason
                    result[PHLDR_TRIP_REASON] = customField.selectedValues[0].label;
                    break;
                case PHLDR_TRANSPORT_TYPE: //Transport type
                    result[PHLDR_TRANSPORT_TYPE] = customField.selectedValues[0].label;
                    break;
                default:
                    break;
            }
        });
    } catch (error) {
        console.log(error);
    }
    return result;
}

function formatShortDate(date: Date): string {
    const options: any = {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    };
    const dateString = date.toLocaleDateString('bg-BG', options)
    return dateString.substring(0, dateString.length - 3);
}

function formatLongDate(date: Date): string {
    const options: any = {
        year: 'numeric',
        month: 'long',
        day: '2-digit',
    };
    const dateString = date.toLocaleDateString('bg-BG', options)
    return dateString.substring(0, dateString.length - 3);
}

async function getEmployeeName(employee: any): Promise<string> {
    return await getEmployeeCyrillicName(employee.id);
}

async function getEmployeeCyrillicName(employeeId: string): Promise<string> {
    const userDetailsResponse = await makePayhawkHttpRequest('GET', `users/${employeeId}/reimbursement-details`, null, null);
    return userDetailsResponse;
}

async function uploadDocumentToPayhawk(expenseId: string, fileBuffer: Buffer, fileName: string): Promise<string> {
    const blob = new Blob(
        [fileBuffer],
        {
            type: "application/pdf"
        }
    );
    const documentResponse = await makePayhawkMultipartFormDataRequest(`expenses/${expenseId}/files`, blob, fileName);
    return documentResponse;
}

async function makePayhawkHttpRequest(
    method: string,
    path: string,
    params: any,
    body: any
): Promise<any> {
    const url = new URL(`${PAYHAWK_API_BASE_URL}/${path}`);

    if (params) {
        Object.keys(params).forEach(key => url.searchParams.append(key, params[key]));
    }

    const options: RequestInit = {
        method: method,
        headers: {
            ...getPayhawkAuthHeader(),
            'Content-Type': 'application/json',
        }
    };
    let urlString = url.toString();
    try {
        const response = await fetch(urlString, options);

        // Check if the response is OK (status code in the range 200-299)
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status} - ${response.statusText}`);
        }

        // Parse the response based on content type (assuming JSON)
        const responseData = await response.json();
        return responseData;
    } catch (error) {
        console.error('Error making Payhawk HTTP request:', error);
        throw error;
    }
}

async function makePayhawkMultipartFormDataRequest(path: string, file: Blob, fileName: string): Promise<any> {
    const form = new FormData();
    form.append('file', file, fileName);

    const headers = getPayhawkAuthHeader();

    const url = `${PAYHAWK_API_BASE_URL}/${path}`;

    const response = await fetch(url, {
        method: 'POST',
        headers: {
            ...headers,
        },
        body: form,
    });

    if (response.ok) {
        console.log('PDF file successfully uploaded to Payhawk');
    }

    return response.json();
}

function getPayhawkAuthHeader(): any {
    return {
        'X-Payhawk-ApiKey': PAYHAWK_API_KEY
    };
};

// Acquire token using client credentials flow
async function getAccessToken(): Promise<string> {
    const tokenRequest = {
        scopes: ['https://graph.microsoft.com/.default'],
    };

    try {
        const cca = new ConfidentialClientApplication(msalConfig);
        const response = await cca.acquireTokenByClientCredential(tokenRequest);
        return response?.accessToken || '';
    } catch (error) {
        console.error('Error acquiring access token:', error);
        throw new Error('Could not acquire access token.');
    }
}

async function copyAndRenameFile(
    siteId: string,
    templateFileId: string,
    realFilesFolderId: string,
    newFileName: string,
    accessToken: string
): Promise<Response> {
    try {
        // Get the file metadata from the Templates folder
        const fileResponse = await fetch(
            `${GRAPH_API_BASE_URL}/drives/${siteId}/items/${templateFileId}`,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                },
            }
        );

        if (!fileResponse.ok) {
            throw new Error(`Error fetching file metadata: ${fileResponse.statusText}`);
        }

        const fileData = await fileResponse.json();
        const fileId = fileData.id;

        // Copy the file to the Real Files folder
        const copyResponse = await fetch(
            `${GRAPH_API_BASE_URL}/drives/${siteId}/items/${templateFileId}/copy`,
            {
                method: 'POST',
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    parentReference: {
                        driveId: siteId,
                        id: realFilesFolderId,
                    },
                    name: newFileName,
                }),
            }
        );

        if (!copyResponse.ok) {
            throw new Error(`Error copying file: ${copyResponse.statusText}`);
        }
        let response = await copyResponse;
        console.log('File copied successfully in Sharepoint');

        return response;
    } catch (error) {
        console.error('Error copying file:', error);
        throw error;
    }
}

const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

async function replacePlaceholdersInFile(
    siteId: string,
    realFilesFolderId: string,
    templateFileId: string,
    placeholders: { [key: string]: string },
    accessToken: string
): Promise<void> {
    try {
        // Get file content as binary
        const fileResponse = await fetch(
            `${GRAPH_API_BASE_URL}/drives/${siteId}/items/${templateFileId}/content`,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                }
            }
        );

        if (!fileResponse.ok) {
            throw new Error(`Error fetching file content: ${fileResponse.statusText}`);
        }

        const fileBuffer = await fileResponse.arrayBuffer();

        // Assuming you are using a library like `docxtemplater` to modify the .docx content.
        let zip = new PizZip(Buffer.from(fileBuffer));
        let doc = new Docxtemplater(zip);

        // Replace placeholders
        doc.setData(placeholders);

        try {
            doc.render();
        } catch (error) {
            console.error('Error rendering document:', error);
            throw error;
        }

        // Generate the updated document
        const updatedDocBuffer = doc.getZip().generate({ type: 'nodebuffer' });

        // Upload the modified file back to the Real Files folder
        const uploadResponse = await fetch(
            `${GRAPH_API_BASE_URL}/drives/${siteId}/items/${templateFileId}/content`,
            {
                method: 'PUT',
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                },
                body: updatedDocBuffer,
            }
        );

        if (!uploadResponse.ok) {
            throw new Error(`Error uploading updated file: ${uploadResponse.statusText}`);
        }

        console.log('File updated successfully with placeholders');
    } catch (error) {
        console.error('Error replacing placeholders in file:', error);
        throw error;
    }
}

async function downloadFileAsPDF(fileid: string, driveId: string, accessToken: string): Promise<Buffer> {
    try {
        // Construct the URL to get the file content
        const fileUrl = `${GRAPH_API_BASE_URL}/drives/${driveId}/items/${fileid}/content?format=pdf`;
        // Fetch the file content from SharePoint
        const response = await fetch(fileUrl, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
            },
        });

        if (!response.ok) {
            throw new Error(`Failed to download file: ${response.statusText}`);
        }

        // Get the file content as a buffer
        const fileBuffer = await response.arrayBuffer();

        console.log('File successfuly downloaded as PDF');

        // Return the file content as a Buffer
        return Buffer.from(fileBuffer);
    } catch (error) {
        console.error(`Error downloading file: ${error.message}`);
        throw error;
    }
}

function extractIdFromUrl(url: string): string | null {
    const regex = /items\/([^\/?]+)\?/;
    const match = url.match(regex);
    return match ? match[1] : null;
}