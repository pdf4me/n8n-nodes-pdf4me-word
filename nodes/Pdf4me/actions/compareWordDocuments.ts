import type { IExecuteFunctions, IDataObject, INodeProperties } from 'n8n-workflow';
import {
	pdf4meAsyncRequest,
	ActionConstants,
} from '../GenericFunctions';

export const description: INodeProperties[] = [
	// === FIRST DOCUMENT INPUT SETTINGS ===
	{
		displayName: 'First Document Input Method',
		name: 'firstInputDataType',
		type: 'options',
		required: true,
		default: 'binaryData',
		description: 'Choose how to provide the first Word file for comparison',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
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
		displayName: 'First Document Binary Property Name',
		name: 'firstBinaryPropertyName',
		type: 'string',
		required: true,
		default: 'data',
		description: 'Name of the binary property containing the first Word file',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
				firstInputDataType: ['binaryData'],
			},
		},
	},
	{
		displayName: 'First Document Base64 Content',
		name: 'firstBase64Content',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		required: true,
		default: '',
		description: 'Base64 encoded string containing the first Word file data',
		placeholder: 'UEsDBAoAAAAAAIdO4kAAAAAAAAAAAAAAAAAJAAAAZG9jUHJvcHMv...',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
				firstInputDataType: ['base64'],
			},
		},
	},
	{
		displayName: 'First Document URL',
		name: 'firstUrl',
		type: 'string',
		required: true,
		default: '',
		description: 'URL to download the first Word file from',
		placeholder: 'https://example.com/original.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
				firstInputDataType: ['url'],
			},
		},
	},
	{
		displayName: 'First Document Name',
		name: 'firstDocName',
		type: 'string',
		default: 'original.docx',
		description: 'Name of the first Word file',
		placeholder: 'original.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	// === SECOND DOCUMENT INPUT SETTINGS ===
	{
		displayName: 'Second Document Input Method',
		name: 'secondInputDataType',
		type: 'options',
		required: true,
		default: 'binaryData',
		description: 'Choose how to provide the second Word file for comparison',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
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
		displayName: 'Second Document Binary Property Name',
		name: 'secondBinaryPropertyName',
		type: 'string',
		required: true,
		default: 'data',
		description: 'Name of the binary property containing the second Word file',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
				secondInputDataType: ['binaryData'],
			},
		},
	},
	{
		displayName: 'Second Document Base64 Content',
		name: 'secondBase64Content',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		required: true,
		default: '',
		description: 'Base64 encoded string containing the second Word file data',
		placeholder: 'UEsDBAoAAAAAAIdO4kAAAAAAAAAAAAAAAAAJAAAAZG9jUHJvcHMv...',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
				secondInputDataType: ['base64'],
			},
		},
	},
	{
		displayName: 'Second Document URL',
		name: 'secondUrl',
		type: 'string',
		required: true,
		default: '',
		description: 'URL to download the second Word file from',
		placeholder: 'https://example.com/revised.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
				secondInputDataType: ['url'],
			},
		},
	},
	{
		displayName: 'Second Document Name',
		name: 'secondDocName',
		type: 'string',
		default: 'revised.docx',
		description: 'Name of the second Word file',
		placeholder: 'revised.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	// === COMPARISON OPTIONS ===
	{
		displayName: 'Ignore Formatting',
		name: 'ignoreFormatting',
		type: 'boolean',
		default: false,
		description: 'Whether to ignore formatting changes when comparing',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	{
		displayName: 'Ignore Case Changes',
		name: 'ignoreCase',
		type: 'boolean',
		default: false,
		description: 'Whether to ignore case changes when comparing',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	{
		displayName: 'Ignore Comments',
		name: 'ignoreComments',
		type: 'boolean',
		default: false,
		description: 'Whether to ignore comments when comparing',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	{
		displayName: 'Ignore Tables',
		name: 'ignoreTables',
		type: 'boolean',
		default: false,
		description: 'Whether to ignore tables when comparing',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	{
		displayName: 'Ignore Fields',
		name: 'ignoreFields',
		type: 'boolean',
		default: false,
		description: 'Whether to ignore fields when comparing',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	{
		displayName: 'Ignore Footnotes',
		name: 'ignoreFootnotes',
		type: 'boolean',
		default: false,
		description: 'Whether to ignore footnotes when comparing',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	{
		displayName: 'Ignore Textboxes',
		name: 'ignoreTextboxes',
		type: 'boolean',
		default: false,
		description: 'Whether to ignore textboxes when comparing',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	{
		displayName: 'Ignore Headers and Footers',
		name: 'ignoreHeadersAndFooters',
		type: 'boolean',
		default: false,
		description: 'Whether to ignore headers and footers when comparing',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	{
		displayName: 'Author',
		name: 'author',
		type: 'string',
		default: 'System Comparison',
		description: 'Author name for the comparison',
		placeholder: 'System Comparison',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'compared_documents.docx',
		description: 'Name for the comparison result Word file',
		placeholder: 'output.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
	{
		displayName: 'Output Binary Data Name',
		name: 'binaryDataName',
		type: 'string',
		default: 'data',
		description: 'Name for the binary data in the n8n output (used to access the processed file)',
		placeholder: 'word-file',
		displayOptions: {
			show: {
				operation: [ActionConstants.CompareWordDocuments],
			},
		},
	},
];

/**
 * Helper function to get document content from different input types
 */
async function getDocumentContent(
	this: IExecuteFunctions,
	index: number,
	inputDataType: string,
	binaryPropertyName: string,
	base64Content: string,
	url: string,
	defaultDocName: string,
): Promise<{ content: string; fileName: string }> {
	let docContent: string;
	let fileName = defaultDocName;

	if (inputDataType === 'binaryData') {
		const item = this.getInputData(index);
		if (!item[0].binary || !item[0].binary[binaryPropertyName]) {
			throw new Error(`No binary data found in property '${binaryPropertyName}'`);
		}
		const binaryData = item[0].binary[binaryPropertyName];
		const buffer = await this.helpers.getBinaryDataBuffer(index, binaryPropertyName);
		docContent = buffer.toString('base64');
		if (binaryData.fileName) {
			fileName = binaryData.fileName;
		}
	} else if (inputDataType === 'base64') {
		docContent = base64Content;
		if (docContent.includes(',')) {
			docContent = docContent.split(',')[1];
		}
	} else if (inputDataType === 'url') {
		if (!url || url.trim() === '') {
			throw new Error('URL is required when using URL input type');
		}
		try {
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
					fileName = filenameMatch[1].replace(/['"]/g, '');
				}
			}
			if (fileName === defaultDocName) {
				const urlParts = url.split('/');
				const urlFilename = urlParts[urlParts.length - 1].split('?')[0];
				if (urlFilename) {
					fileName = decodeURIComponent(urlFilename);
				}
			}
		} catch (error) {
			const errorMessage = error instanceof Error ? error.message : 'Unknown error';
			throw new Error(`Failed to download file from URL: ${errorMessage}`);
		}
	} else {
		throw new Error(`Unsupported input data type: ${inputDataType}`);
	}

	if (!docContent || docContent.trim() === '') {
		throw new Error('Document content is required');
	}

	return { content: docContent, fileName };
}

/**
 * Compare two Word documents using PDF4Me API
 * Process: Read both Word files → Encode to base64 → Send API request → Poll for completion → Save comparison result
 * Compares two Word documents and creates a marked-up document showing differences
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;
		const firstDocName = this.getNodeParameter('firstDocName', index) as string;
		const secondDocName = this.getNodeParameter('secondDocName', index) as string;

		// Get comparison options
		const ignoreFormatting = this.getNodeParameter('ignoreFormatting', index, false) as boolean;
		const ignoreCase = this.getNodeParameter('ignoreCase', index, false) as boolean;
		const ignoreComments = this.getNodeParameter('ignoreComments', index, false) as boolean;
		const ignoreTables = this.getNodeParameter('ignoreTables', index, false) as boolean;
		const ignoreFields = this.getNodeParameter('ignoreFields', index, false) as boolean;
		const ignoreFootnotes = this.getNodeParameter('ignoreFootnotes', index, false) as boolean;
		const ignoreTextboxes = this.getNodeParameter('ignoreTextboxes', index, false) as boolean;
		const ignoreHeadersAndFooters = this.getNodeParameter('ignoreHeadersAndFooters', index, false) as boolean;
		const author = this.getNodeParameter('author', index, 'System Comparison') as string;

		// Get first document
		const firstInputDataType = this.getNodeParameter('firstInputDataType', index) as string;
		const firstBinaryPropertyName = this.getNodeParameter('firstBinaryPropertyName', index) as string;
		const firstBase64Content = this.getNodeParameter('firstBase64Content', index) as string;
		const firstUrl = this.getNodeParameter('firstUrl', index) as string;

		const firstDoc = await getDocumentContent.call(
			this,
			index,
			firstInputDataType,
			firstBinaryPropertyName,
			firstBase64Content,
			firstUrl,
			firstDocName,
		);

		// Get second document
		const secondInputDataType = this.getNodeParameter('secondInputDataType', index) as string;
		const secondBinaryPropertyName = this.getNodeParameter('secondBinaryPropertyName', index) as string;
		const secondBase64Content = this.getNodeParameter('secondBase64Content', index) as string;
		const secondUrl = this.getNodeParameter('secondUrl', index) as string;

		const secondDoc = await getDocumentContent.call(
			this,
			index,
			secondInputDataType,
			secondBinaryPropertyName,
			secondBase64Content,
			secondUrl,
			secondDocName,
		);

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				Name: firstDoc.fileName,
			},
			docContent: firstDoc.content,
			compareWith: secondDoc.content,
			ignoreFormatting: ignoreFormatting,
			ignoreCase: ignoreCase,
			ignoreComments: ignoreComments,
			ignoreTables: ignoreTables,
			ignoreFields: ignoreFields,
			ignoreFootnotes: ignoreFootnotes,
			ignoreTextboxes: ignoreTextboxes,
			ignoreHeadersAndFooters: ignoreHeadersAndFooters,
			author,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Word/CompareDocuments',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = firstDoc.fileName
					? firstDoc.fileName.replace(/\.[一个人在.]*$/, '')
					: 'compared_documents';
				fileName = `${baseName}_compared.docx`;
			}

			// Ensure .docx extension
			if (!fileName.toLowerCase().endsWith('.docx')) {
				fileName = `${fileName.replace(/\.[^.]*$/, '')}.docx`;
			}

			// Handle the response - Word API returns JSON with embedded base64 file
			let wordBuffer: Buffer;

			// Check for Buffer first to properly narrow TypeScript types
			if (Buffer.isBuffer(responseData)) {
				wordBuffer = responseData;
			} else if (typeof responseData === 'string') {
				wordBuffer = Buffer.from(responseData, 'base64');
			} else if (typeof responseData === 'object' && responseData !== null) {
				const response = responseData as IDataObject;

				if (response.document) {
					const document = response.document;
					if (typeof document === 'string') {
						wordBuffer = Buffer.from(document, 'base64');
					} else if (typeof document === 'object' && document !== null) {
						const docObj = document as IDataObject;
						const docContent =
							(docObj.docData as string) ||
							(docObj.content as string) ||
							(docObj.docContent as string) ||
							(docObj.data as string) ||
							(docObj.file as string);

						if (!docContent) {
							const docKeys = Object.keys(docObj).join(', ');
							throw new Error(`Document object has unexpected structure. Available keys: ${docKeys}`);
						}
						wordBuffer = Buffer.from(docContent, 'base64');
					} else {
						throw new Error(`Document field is neither string nor object: ${typeof document}`);
					}
				} else {
					const docContent =
						(response.docData as string) ||
						(response.content as string) ||
						(response.fileContent as string) ||
						(response.data as string);

					if (!docContent) {
						const keys = Object.keys(responseData).join(', ');
						throw new Error(`Word API returned unexpected JSON structure. Available keys: ${keys}`);
					}
					wordBuffer = Buffer.from(docContent, 'base64');
				}
			} else {
				throw new Error(`Unexpected response format: ${typeof responseData}`);
			}

			// Validate the response contains Word data
			if (!wordBuffer || wordBuffer.length < 1000) {
				throw new Error(
					'Invalid Word response from API. The file appears to be too small or corrupted.',
				);
			}

			// Validate Word file format (DOCX files start with PK signature - ZIP format)
			const magicBytes = wordBuffer.toString('hex', 0, 4);
			if (magicBytes !== '504b0304') {
				throw new Error(
					`Invalid Word file format. Expected DOCX file but got unexpected data. Magic bytes: ${magicBytes}`,
				);
			}

			// Create binary data for output
			const binaryData = await this.helpers.prepareBinaryData(
				wordBuffer,
				fileName,
				'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			);

			// Determine the binary data name
			const binaryDataKey = binaryDataName || 'data';

			return [
				{
					json: {
						fileName,
						fileSize: wordBuffer.length,
						success: true,
						firstDocument: firstDoc.fileName,
						secondDocument: secondDoc.fileName,
						comparisonOptions: {
							ignoreFormatting,
							ignoreCase,
							ignoreComments,
							ignoreTables,
							ignoreFields,
							ignoreFootnotes,
							ignoreTextboxes,
							ignoreHeadersAndFooters,
							author,
						},
						message: 'Word documents compared successfully',
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
		throw new Error(`Compare Word documents failed: ${errorMessage}`);
	}
}

