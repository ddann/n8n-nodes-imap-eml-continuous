import { FetchQueryObject, ImapFlow } from "imapflow";
import { IBinaryKeyData, IDataObject, IExecuteFunctions, INodeExecutionData } from "n8n-workflow";
import { IResourceOperationDef } from "../../../utils/CommonDefinitions";
import { getMailboxPathFromNodeParameter, parameterSelectMailbox } from "../../../utils/SearchFieldParameters";
import { emailSearchParameters } from "../../../utils/EmailSearchParameters";
import { getEmailSearchParametersFromNode } from "../../../utils/EmailSearchParameters";

// Copy EmailParts enum from Get Many for use in options
enum EmailParts {
  BodyStructure = 'bodyStructure',
  Flags = 'flags',
  Size = 'size',
  AttachmentsInfo = 'attachmentsInfo',
  TextContent = 'textContent',
  HtmlContent = 'htmlContent',
  Headers = 'headers',
}

export const downloadOperation: IResourceOperationDef = {
  operation: {
    name: 'Download as EML',
    value: 'downloadEml',
  },
  parameters: [
    {
      ...parameterSelectMailbox,
      description: 'Select the mailbox',
    },
    ...emailSearchParameters,
    {
      displayName: 'Include Message Parts',
      name: 'includeParts',
      type: 'multiOptions',
      placeholder: 'Add Part',
      default: [],
      options: [
        {
          name: 'Text Content',
          value: EmailParts.TextContent,
        },
        {
          name: 'HTML Content',
          value: EmailParts.HtmlContent,
        },
        {
          name: 'Attachments Info',
          value: EmailParts.AttachmentsInfo,
        },
        {
          name: 'Flags',
          value: EmailParts.Flags,
        },
        {
          name: 'Size',
          value: EmailParts.Size,
        },
        {
          name: 'Body Structure',
          value: EmailParts.BodyStructure,
        },
        {
          name: 'Headers',
          value: EmailParts.Headers,
        },
      ],
    },
    {
      displayName: 'Include All Headers',
      name: 'includeAllHeaders',
      type: 'boolean',
      default: true,
      description: 'Whether to include all headers in the output',
      displayOptions: {
        show: {
          includeParts: [
            EmailParts.Headers,
          ],
        },
      },
    },
    {
      displayName: 'Headers to Include',
      name: 'headersToInclude',
      type: 'string',
      default: '',
      description: 'Comma-separated list of headers to include',
      placeholder: 'received,authentication-results,return-path',
      displayOptions: {
        show: {
          includeParts: [
            EmailParts.Headers,
          ],
          includeAllHeaders: [
            false,
          ],
        },
      },
    },
    {
      displayName: 'Email UID',
      name: 'emailUid',
      type: 'string',
      default: '',
      description: 'UID of the email to download',
    },
    {
      displayName: 'Output to Binary Data',
      name: 'outputToBinary',
      type: 'boolean',
      default: true,
      description: 'Whether to output the email as binary data or JSON as text',
      hint: 'If true, the email will be output as binary data. If false, the email will be output as JSON as text.',
    },
    {
      displayName: 'Put Output File in Field',
      name: 'binaryPropertyName',
      type: 'string',
      default: 'data',
      required: true,
      placeholder: 'e.g data',
      hint: 'The name of the output binary field to put the file in',
      displayOptions: {
        show: {
          outputToBinary: [true],
        },
      },
    },
  ],
  async executeImapAction(context: IExecuteFunctions, itemIndex: number, client: ImapFlow): Promise<INodeExecutionData[] | null> {
    const mailboxPath = getMailboxPathFromNodeParameter(context, itemIndex);
    await client.mailboxOpen(mailboxPath, { readOnly: true });

    const emailUid = context.getNodeParameter('emailUid', itemIndex) as string;
    const outputToBinary = context.getNodeParameter('outputToBinary', itemIndex, true) as boolean;
    const binaryPropertyName = context.getNodeParameter('binaryPropertyName', itemIndex, 'data',) as string;

    const query: FetchQueryObject = {
      uid: true,
      source: true,
    };

    let results: INodeExecutionData[] = [];

    if (emailUid) {
      // Download a single email by UID (current behavior)
      const emailInfo = await client.fetchOne(emailUid, query, { uid: true });
      if (!emailInfo || !emailInfo.source) {
        throw new Error('No email found with the specified UID');
      }
      let binaryFields: IBinaryKeyData | undefined = undefined;
      let jsonData: IDataObject = {};
      if (outputToBinary) {
        const binaryData = await context.helpers.prepareBinaryData(emailInfo.source, mailboxPath + '_' + emailUid + '.eml', 'message/rfc822');
        binaryFields = {
          [binaryPropertyName]: binaryData,
        };
      } else {
        jsonData = {
          ...jsonData,
          emlContent: emailInfo.source.toString(),
        };
      }
      results.push({
        json: jsonData,
        binary: binaryFields,
        pairedItem: { item: itemIndex },
      });
    } else {
      // Download all emails matching search parameters (or all in mailbox)
      const searchObject = getEmailSearchParametersFromNode(context, itemIndex);
      let foundAny = false;
      for await (const email of client.fetch(searchObject, query)) {
        foundAny = true;
        let binaryFields: IBinaryKeyData | undefined = undefined;
        let jsonData: IDataObject = {};
        if (outputToBinary) {
          const binaryData = await context.helpers.prepareBinaryData(email.source, mailboxPath + '_' + email.uid + '.eml', 'message/rfc822');
          binaryFields = {
            [binaryPropertyName]: binaryData,
          };
        } else {
          jsonData = {
            ...jsonData,
            emlContent: email.source.toString(),
          };
        }
        results.push({
          json: jsonData,
          binary: binaryFields,
          pairedItem: { item: itemIndex },
        });
      }
      if (!foundAny) {
        throw new Error('No emails found matching the search criteria.');
      }
    }
    return results;
  },
};
