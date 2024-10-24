import axios, { AxiosResponse } from 'axios';
var toArray = require('stream-to-array');

import * as fs from 'fs';
import * as path from 'path';
import { ConfidentialClientApplication } from '@azure/msal-node';

const ACCOUNT_ID = "";
const PAYHAWK_API_KEY = "";
const WEBHOOK_EVENT_NAME = "expense.reviewed";
const PAYHAWK_API_BASE_URL = `https://api.payhawk.com/api/v3/accounts/${ACCOUNT_ID}`;

const driveId = '';
const templateFileId = '';
const realFilesFolderId = '';

const msalConfig = {
  auth: {
    clientId: '', // Replace with your client ID
    authority: 'https://login.microsoftonline.com/', // Replace with your tenant ID
    clientSecret: '' // Replace with your client secret
  }
};

const folderUrl = ''; // SharePoint folder URL
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
const PHLDR_TRIP_REASON = 'trip_reason';                    //Value from custom field - Причина за командировка
const PHLDR_TRANSPORT_TYPE = 'transport_type';              //Value from custom field - Вид транспортно средство
const PHLDR_WORK_TITLE = 'work_title';                      //Value from custom field - Длъжност
const PHLDR_TRIP_TOTAL_AMOUNT = 'trip_total_amount';        //The total amount and currency for the trip


// Initialize MSAL client
const cca = new ConfidentialClientApplication(msalConfig);

export class PerDiemDocumentBuilder {

  async initWebhook(callbackUrl: string) {

    try {
      const webhookEventType = process.env.WEBHOOK_EVENT_NAME;
      const webhooksResponse = await this.makePayhawkHttpRequest('GET', '/webhooks', null, null);
      const webhooks = webhooksResponse.data.items;
      const existingWebhook = webhooks.find((webhook: any) => webhook.callbackUrl === callbackUrl && webhook.eventType === webhookEventType);
      if (existingWebhook) {
        console.log('Webhook already exists!');
      } else {

      }
      return webhooksResponse.data.items;
    } catch (error) {
      console.log(error);
    }
  }

  async generatePerDiemDocument(expenseId: string, regenerateIfExists: boolean): Promise<string> {
    try {
      // Get expense

      const expense = await this.getExpense(expenseId);

      // If the expense is not of type perDiem - abort
      if (expense.type !== 'perDiem') {
        return 'Expense is not a per-diem.';
      }

      if (!regenerateIfExists && expense.document.files.length > 1) {
        return 'Custom per-diem form already generated for this expense.';
      }

      // Get expense data
      const expenseData = await this.getExpenseData(expense);

      const accessToken = await this.getAccessToken();

      const newFileName = `Per Diem For ${expenseId}.docx`;
      console.log(newFileName);
      // Copy and rename the file
      const newFile = await this.copyAndRenameFile(
        driveId,
        templateFileId,
        realFilesFolderId,
        newFileName,
        accessToken,
      );

      // Replace placeholders in the copied file
      const placeholders = {
        expense_id: expense.id,
        expense_created_date: expense.createdAt,
        employee_name: expense.createdBy,
        employee_parent_team: expenseData.PHLDR_EMPLOYEE_TEAM,
        destination: expenseData.PHLDR_DESTINATION,
        trip_reason: expenseData.PHLDR_TRIP_REASON,
        from_date: expenseData.PHLDR_FROM_DATE,
        to_date: expenseData.PHLDR_TO_DATE,
        approver_name: expenseData.PHLDR_APPROVER_NAME,
      };

      await this.replacePlaceholdersInFile(
        driveId,
        realFilesFolderId,
        templateFileId,
        placeholders,
        accessToken
      );

      //const newFileId = await this.copyTamplateFile(expenseData);

      //Replace placeholders in the copied template with the correct data
      //await this.replaceTextInDoc(newFileId, expenseData);

      //Get the file as PDF
      //const pdf:any = await this.exportFileAsPdf(newFileId);
      //const arr = await toArray(pdf.data);
      //const buffers = arr.map(part:any => util.isBuffer(part) ? part : Buffer.from(part));
      //const b:Buffer = Buffer.concat(arr);

      //Attach the document to the expense in Payhawk
      //const documentResponse = await this.uploadDocumentToPayhawk(expenseId, b);

      return 'Success';
    } catch (error) {
      console.error(error);
      return 'Error'
    }
  }

  private async copyTamplateFile(expenseData: any): Promise<string> {
    throw new Error('Unable to duplicate template file: ');
    // https://learn.microsoft.com/en-us/graph/api/driveitem-copy?view=graph-rest-1.0&tabs=http
  }

  private async replaceTextInDoc(copiedTemplateFileId: string, replaceObject: any): Promise<void> {
  }

  private async exportFileAsPdf(googleDriveFileId: string): Promise<void> {
    // https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format?view=graph-rest-1.0&tabs=http
  }


  private async getExpense(expenseId: string): Promise<any> {
    try {
      const expenseResponse = await this.makePayhawkHttpRequest('GET', `expenses/${expenseId}`, null, null);
      return expenseResponse.data;
    } catch (error) {
      console.log(error);
    }
  }

  private async getExpenseData(expense: any): Promise<any> {
    const expenseId = expense.id;
    const result: any = {};
    try {
      const expenseWorkflowResponse = await this.makePayhawkHttpRequest('GET', `expenses/${expenseId}/workflow`, null, null);
      const expenseWorkflow = expenseWorkflowResponse.data;

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
      result[PHLDR_FROM_DATE] = this.formatShortDate(new Date(Date.parse(firstStop.date)));
      result[PHLDR_TO_DATE] = this.formatShortDate(new Date(Date.parse(lastStop.date)));

      //Expense created date
      const createdAt: Date = new Date(Date.parse(expense.createdAt));
      result[PHLDR_EXPENSE_DATE] = this.formatLongDate(createdAt);

      //Employee name
      result[PHLDR_EMPLOYEE_NAME] = await this.getEmployeeName(expense.createdBy);

      //Trip total amount
      var totalAmountFormat = new Intl.NumberFormat('en-US', {
        minimumIntegerDigits: 1,
        minimumFractionDigits: 2,
        useGrouping: false
      });
      result[PHLDR_TRIP_TOTAL_AMOUNT] = totalAmountFormat.format(expense.reconciliation.totalAmount);


      //Approver name
      if (expenseWorkflow.approvedBy) {
        result[PHLDR_APPROVER_NAME] = await this.getEmployeeName(expenseWorkflow.approvedBy);
      } else {
        result[PHLDR_APPROVER_NAME] = 'not approved yet';
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

  private formatShortDate(date: Date): string {
    const options: any = {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    };
    const dateString = date.toLocaleDateString('bg-BG', options)
    return dateString.substring(0, dateString.length - 3);
  }

  private formatLongDate(date: Date): string {
    const options: any = {
      year: 'numeric',
      month: 'long',
      day: '2-digit',
    };
    const dateString = date.toLocaleDateString('bg-BG', options)
    return dateString.substring(0, dateString.length - 3);
  }

  private async getEmployeeName(employee: any): Promise<string> {
    return await this.getEmployeeCyrillicName(employee.id);
  }

  private async getEmployeeCyrillicName(employeeId: string): Promise<string> {
    const userDetailsResponse = await this.makePayhawkHttpRequest('GET', `users/${employeeId}/reimbursement-details`, null, null);
    const userDetails = userDetailsResponse.data;
    return userDetails.accountHolder;
  }

  private async uploadDocumentToPayhawk(expenseId: string, fileBuffer: Buffer): Promise<string> {
    const blob = new Blob(
      [fileBuffer],
      {
        type: "application/pdf"
      }
    );
    const userDetailsResponse = await this.makePayhawkMultipartFormDataRequest(`expenses/${expenseId}/files`, blob);
    const userDetails = userDetailsResponse.data;
    return userDetails.accountHolder;
  }

  private async makePayhawkHttpRequest(method: string, path: string, params: any, body: any): Promise<AxiosResponse<any, any>> {
    const url = `${PAYHAWK_API_BASE_URL}/${path}`
    const response = await axios.request({
      method: method,
      url: url,
      headers: this.getPayhawkAuthHeader(),
      data: body,
      params: params
    });

    return response;
  }

  private async makePayhawkMultipartFormDataRequest(path: string, file: Blob): Promise<AxiosResponse<any, any>> {
    const form = new FormData();
    form.append('my_buffer.pdf', file);

    const headers = this.getPayhawkAuthHeader();
    headers['Content-Type'] = 'multipart/form-data';

    const url = `${PAYHAWK_API_BASE_URL}/${path}`;
    const response = axios.post(
      url,
      form,
      { headers }
    )

    return response;
  }

  private getPayhawkAuthHeader(): any {
    return {
      'X-Payhawk-ApiKey': PAYHAWK_API_KEY
    };
  };

  // Acquire token using client credentials flow
  private async getAccessToken(): Promise<string> {
    const tokenRequest = {
      scopes: ['https://graph.microsoft.com/.default'],
    };

    try {
      const response = await cca.acquireTokenByClientCredential(tokenRequest);
      return response?.accessToken || '';
    } catch (error) {
      console.error('Error acquiring access token:', error);
      throw new Error('Could not acquire access token.');
    }
  }

  // Download a file from SharePoint
  private async downloadFile(fileName: string) {
    const accessToken = await this.getAccessToken();
    const fileEndpoint = `${folderUrl}/${fileName}`;

    try {
      const response = await axios.get(fileEndpoint, {
        headers: {
          Authorization: `Bearer ${accessToken}`
        },
        responseType: 'stream'
      });

      const filePath = path.join(__dirname, fileName);
      const writer = fs.createWriteStream(filePath);

      response.data.pipe(writer);

      writer.on('finish', () => {
        console.log(`File downloaded successfully: ${filePath}`);
      });

      writer.on('error', (err) => {
        console.error('Error writing file:', err);
      });
    } catch (error) {
      console.error('Error downloading file:', error);
    }
  }

  // Upload a file to SharePoint
  private async uploadFile(filePath: string) {
    const accessToken = await this.getAccessToken();
    const fileName = path.basename(filePath);
    const fileEndpoint = `${folderUrl}/${fileName}`;

    try {
      const fileStream = fs.createReadStream(filePath);
      const response = await axios.put(fileEndpoint, fileStream, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/octet-stream'
        }
      });

      if (response.status === 201) {
        console.log(`File uploaded successfully: ${fileName}`);
      } else {
        console.error('Error uploading file:', response.status, response.statusText);
      }
    } catch (error) {
      console.error('Error uploading file:', error);
    }
  }

  async copyAndRenameFile(
    siteId: string,
    templateFileId: string,
    realFilesFolderId: string,
    newFileName: string,
    accessToken: string
  ): Promise<string> {
    try {
      // Get the file metadata from the Templates folder
      const fileResponse = await axios.get(
        `${GRAPH_API_BASE_URL}/drives/${siteId}/items/${templateFileId}`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const fileId = fileResponse.data.id;

      // Copy the file to the Real Files folder
      const copyResponse = await axios.post(
        `${GRAPH_API_BASE_URL}/drives/${siteId}/items/${templateFileId}/copy`,
        {
          parentReference: {
            driveId: siteId,
            id: realFilesFolderId,
          },
          name: newFileName,
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      console.log('File copied successfully:', copyResponse);

      return newFileName;
    } catch (error) {
      console.error('Error copying file:', error);
      throw error;
    }
  }

  // Function to replace placeholders in a .docx file with given values
  async replacePlaceholdersInFile(
    siteId: string,
    realFilesFolderId: string,
    templateFileId: string,
    placeholders: { [key: string]: string },
    accessToken: string
  ): Promise<void> {
    try {
      // Get file content as binary
      const fileResponse = await axios.get(
        `${GRAPH_API_BASE_URL}/drives/${siteId}/items/${templateFileId}/content`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          },
          responseType: 'arraybuffer',
        }
      );

      let fileBuffer = Buffer.from(fileResponse.data);

      // Assuming you are using a library like `docxtemplater` to modify the .docx content.
      // You will need to install and import it.
      const PizZip = require('pizzip');
      const Docxtemplater = require('docxtemplater');

      let zip = new PizZip(fileBuffer);
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
      await axios.put(
        `${GRAPH_API_BASE_URL}/drives/${siteId}/items/${templateFileId}/content`,
        updatedDocBuffer,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          },
        }
      );

      console.log('File updated successfully:', templateFileId);
    } catch (error) {
      console.error('Error replacing placeholders in file:', error);
      throw error;
    }
  }
}