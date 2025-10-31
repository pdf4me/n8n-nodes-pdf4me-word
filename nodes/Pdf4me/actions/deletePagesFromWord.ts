import type { IExecuteFunctions, IDataObject, INodeExecutionData, INodeProperties } from 'n8n-workflow';
import {
	pdf4meAsyncRequest,
	ActionConstants,
} from '../GenericFunctions';

export const description: INodeProperties[] = [
	// === INPUT FILE SETTINGS ===
	{
		displayName: 'Word File Input Method',
		name: 'inputDataType',
		type: 'options',
		required: true,
		default: 'binaryData',
		description: 'Choose how to provide the Word file for processing',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
			},
		},
		options: [
			{
				name: 'From Previous Node (Binary Data)',
				value: 'binaryData',
				description: 'Use Word file passed from a previous n8n node',
			},
			{
				name: 'Base64 Encoded String',
				value: 'base64',
				description: 'Provide Word file content as base64 encoded string',
			},
			{
				name: 'Download from URL',
				value: 'url',
				description: 'Download Word file directly from a web URL',
			},
		],
	},
	{
		displayName: 'Binary Data Property Name',
		name: 'binaryPropertyName',
		type: 'string',
		required: true,
		default: 'data',
		description: 'Name of the binary property containing the Word file (usually \'data\')',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
				inputDataType: ['binaryData'],
			},
		},
	},
	{
		displayName: 'Base64 Encoded Word Content',
		name: 'base64Content',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		required: true,
		default: '',
		description: 'Base64 encoded string containing the Word file data',
		placeholder: 'UEsDBAoAAAAAAIdO4kAAAAAAAAAAAAAAAAAJAAAAZG9jUHJvcHMv...',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
				inputDataType: ['base64'],
			},
		},
	},
	{
		displayName: 'Word File URL',
		name: 'url',
		type: 'string',
		required: true,
		default: '',
		description: 'URL to download the Word file from (must be publicly accessible)',
		placeholder: 'https://example.com/file.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
				inputDataType: ['url'],
			},
		},
	},
	// === DELETE PAGES SETTINGS ===
	{
		displayName: 'Start Page',
		name: 'startPage',
		type: 'number',
		default: 0,
		description: 'Page number to begin deleting pages from (1-based index, 0 = not used)',
		placeholder: '1',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
			},
		},
	},
	{
		displayName: 'End Page',
		name: 'endPage',
		type: 'number',
		default: 0,
		description: 'Page number to stop deleting pages on (0 = defaults to last page of document)',
		placeholder: '10',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
			},
		},
	},
	{
		displayName: 'Page Numbers',
		name: 'pageNumbers',
		type: 'string',
		default: '',
		description: 'Comma-separated list of page numbers to delete (e.g., "1,3,4"). Can be used alone or combined with Start/End Page',
		placeholder: '1,3,4',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
			},
		},
	},
	// === DOCUMENT SETTINGS ===
	{
		displayName: 'Source Document Name',
		name: 'docName',
		type: 'string',
		default: 'document.docx',
		description: 'Name of the original Word file (for reference and processing)',
		placeholder: 'document.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
			},
		},
	},
	{
		displayName: 'Culture Name',
		name: 'cultureName',
		type: 'string',
		default: 'en-US',
		description: 'Culture name for document (e.g., en-US, fr-FR, de-DE)',
		placeholder: 'en-US',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'document_pages_deleted.docx',
		description: 'Name for the resulting Word file',
		placeholder: 'document_pages_deleted.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
			},
		},
	},
	{
		displayName: 'Output Binary Data Name',
		name: 'binaryDataName',
		type: 'string',
		default: 'data',
		description: 'Name for the binary data in the n8n output',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.DeletePagesFromWord],
			},
		},
	},
];

/**
 * Delete pages from a Word document using PDF4Me API
 */
export async function execute(this: IExecuteFunctions, index: number): Promise<INodeExecutionData[]> {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;
		const outputFileNameParam = this.getNodeParameter('outputFileName', index) as string;
		const startPage = this.getNodeParameter('startPage', index, 0) as number;
		const endPage = this.getNodeParameter('endPage', index, 0) as number;
		const pageNumbers = this.getNodeParameter('pageNumbers', index, '') as string;
		const cultureName = this.getNodeParameter('cultureName', index, 'en-US') as string;

		let docContent: string;
		let originalFileName = docName;

		if (inputDataType === 'binaryData') {
			const binaryPropertyName = this.getNodeParameter('binaryPropertyName', index) as string;
			const item = this.getInputData(index);
			if (!item[0].binary || !item[0].binary[binaryPropertyName]) {
				throw new Error(`No binary data found in property '${binaryPropertyName}'`);
			}
			const binaryData = item[0].binary[binaryPropertyName];
			const buffer = await this.helpers.getBinaryDataBuffer(index, binaryPropertyName);
			docContent = buffer.toString('base64');
			if (binaryData.fileName) {
				originalFileName = binaryData.fileName;
			}
		} else if (inputDataType === 'base64') {
			docContent = this.getNodeParameter('base64Content', index) as string;
			if (docContent.includes(',')) {
				docContent = docContent.split(',')[1];
			}
		} else if (inputDataType === 'url') {
			const url = this.getNodeParameter('url', index) as string;
			if (!url || url.trim() === '') {
				throw new Error('URL is required when using URL input type');
			}
			const response = await this.helpers.httpRequest({
				method: 'GET',
				url,
				encoding: 'arraybuffer',
				returnFullResponse: true,
			});
			const buffer = Buffer.from(response.body as ArrayBuffer);
			docContent = buffer.toString('base64');
			const contentDisposition = response.headers['content-disposition'];
			if (contentDisposition) {
				const filenameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
				if (filenameMatch && filenameMatch[1]) {
					originalFileName = filenameMatch[1].replace(/['"]/g, '');
				}
			}
			if (originalFileName === docName) {
				const urlParts = url.split('/');
				const urlFilename = urlParts[urlParts.length - 1].split('?')[0];
				if (urlFilename) {
					originalFileName = decodeURIComponent(urlFilename);
				}
			}
		} else {
			throw new Error(`Unsupported input data type: ${inputDataType}`);
		}

		if (!docContent || docContent.trim() === '') {
			throw new Error('Word content is required');
		}

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				name: originalFileName,
			},
			docContent,
		};

		// Add StartPage if provided (non-zero)
		if (startPage && startPage > 0) {
			body.StartPage = startPage;
		}

		// Add EndPage if provided (non-zero)
		if (endPage && endPage > 0) {
			body.EndPage = endPage;
		}

		// Add PageNumbers if provided
		if (pageNumbers && pageNumbers.trim() !== '') {
			body.PageNumbers = pageNumbers.trim();
		}

		// Add CultureName if provided
		if (cultureName && cultureName.trim() !== '') {
			body.CultureName = cultureName;
		}

		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Word/DeletePages',
			body,
		);

		if (responseData) {
			let fileName = outputFileNameParam || 'document_pages_deleted.docx';
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName ? originalFileName.replace(/\.[^.]*$/, '') : 'document';
				fileName = `${baseName}_pages_deleted.docx`;
			}
			if (!fileName.toLowerCase().endsWith('.docx')) {
				fileName = `${fileName.replace(/\.[^.]*$/, '')}.docx`;
			}

			let wordBuffer: Buffer;
			if (Buffer.isBuffer(responseData)) {
				wordBuffer = responseData;
			} else if (typeof responseData === 'string') {
				wordBuffer = Buffer.from(responseData, 'base64');
			} else if (typeof responseData === 'object' && responseData !== null) {
				const response = responseData as IDataObject;
				if (response.document) {
					const document = response.document as unknown;
					if (typeof document === 'string') {
						wordBuffer = Buffer.from(document, 'base64');
					} else if (typeof document === 'object' && document !== null) {
						const docObj = document as IDataObject;
						const docContentResp =
							(docObj.docData as string) ||
							(docObj.content as string) ||
							(docObj.docContent as string) ||
							(docObj.data as string) ||
							(docObj.file as string);
						if (!docContentResp) {
							const docKeys = Object.keys(docObj).join(', ');
							throw new Error(`Document object has unexpected structure. Available keys: ${docKeys}`);
						}
						wordBuffer = Buffer.from(docContentResp, 'base64');
					} else {
						throw new Error(`Document field is neither string nor object: ${typeof document}`);
					}
				} else {
					const docContentResp =
						(response.docData as string) ||
						(response.content as string) ||
						(response.fileContent as string) ||
						(response.data as string);
					if (!docContentResp) {
						const keys = Object.keys(responseData).join(', ');
						throw new Error(`Word API returned unexpected JSON structure. Available keys: ${keys}`);
					}
					wordBuffer = Buffer.from(docContentResp, 'base64');
				}
			} else {
				throw new Error(`Unexpected response format: ${typeof responseData}`);
			}

			if (!wordBuffer || wordBuffer.length < 1000) {
				throw new Error('Invalid Word response from API. The file appears to be too small or corrupted.');
			}

			const magicBytes = wordBuffer.toString('hex', 0, 4);
			if (magicBytes !== '504b0304') {
				throw new Error(`Invalid Word file format. Expected DOCX file but got unexpected data. Magic bytes: ${magicBytes}`);
			}

			const binaryData = await this.helpers.prepareBinaryData(
				wordBuffer,
				fileName,
				'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			);

			const binaryDataKey = binaryDataName || 'data';

			return [
				{
					json: {
						fileName,
						fileSize: wordBuffer.length,
						success: true,
						originalFileName,
						startPage: startPage || undefined,
						endPage: endPage || undefined,
						pageNumbers: pageNumbers || undefined,
						cultureName,
						message: 'Successfully deleted pages from Word document',
					},
					binary: {
						[binaryDataKey]: binaryData,
					},
				},
			];
		}

		throw new Error('No response data received from PDF4ME API');
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
		throw new Error(`Delete pages from Word failed: ${errorMessage}`);
	}
}



