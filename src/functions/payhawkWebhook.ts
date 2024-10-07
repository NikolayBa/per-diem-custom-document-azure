import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { PerDiemDocumentBuilder } from "../PerdiemDocumentBuilder";

export async function payhawkWebhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log(`Http function processed request for url "${request.url}"`);

    let bodyJson = await request.json();
    let bodyString = JSON.stringify(bodyJson);

    const perDiemDocumentBuilder = new PerDiemDocumentBuilder();
    let result = await perDiemDocumentBuilder.generatePerDiemDocument(bodyJson["payload"]["expenseId"], true);

    return { body: `${`Received: ${bodyString}`}!` };
};

app.http('payhawkWebhook', {
    methods: ['GET', 'POST'],
    authLevel: 'anonymous',
    handler: payhawkWebhook
});
