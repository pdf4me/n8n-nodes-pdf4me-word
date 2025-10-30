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
				operation: [ActionConstants.UpdateHeadersFooters],
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
				operation: [ActionConstants.UpdateHeadersFooters],
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
				operation: [ActionConstants.UpdateHeadersFooters],
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
				operation: [ActionConstants.UpdateHeadersFooters],
				inputDataType: ['url'],
			},
		},
	},
	// === HEADERS AND FOOTERS SETTINGS ===
	{
		displayName: 'Content Type',
		name: 'contentType',
		type: 'options',
		default: 'plain',
		description: 'Type of content to use for headers and footers',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
			},
		},
		options: [
			{
				name: 'Plain Text',
				value: 'plain',
			},
			{
				name: 'HTML',
				value: 'html',
			},
		],
	},
	{
		displayName: 'All Pages Header',
		name: 'allPagesHeader',
		type: 'string',
		default: '',
		description: 'Content for header on all pages',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['plain'],
			},
		},
		placeholder: 'Document Title',
	},
	{
		displayName: 'All Pages Footer',
		name: 'allPagesFooter',
		type: 'string',
		default: '',
		description: 'Content for footer on all pages',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['plain'],
			},
		},
		placeholder: 'Page {PAGE} of {NUMPAGES}',
	},
	{
		displayName: 'All Pages Header (HTML)',
		name: 'allPagesHeaderHtml',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		default: '',
		description: 'HTML content for header on all pages',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['html'],
			},
		},
		placeholder: '<div style="text-align: center;">Document Header</div>',
	},
	{
		displayName: 'All Pages Footer (HTML)',
		name: 'allPagesFooterHtml',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		default: '',
		description: 'HTML content for footer on all pages',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['html'],
			},
		},
		placeholder: '<div style="text-align: center;">Page {PAGE} of {NUMPAGES}</div>',
	},
	{
		displayName: 'First Page Header (HTML)',
		name: 'firstPageHeaderHtml',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		default: '',
		description: 'HTML content for header on first page only',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['html'],
			},
		},
		placeholder: '<div>Cover Page Header</div>',
	},
	{
		displayName: 'First Page Footer (HTML)',
		name: 'firstPageFooterHtml',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		default: '',
		description: 'HTML content for footer on first page only',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['html'],
			},
		},
		placeholder: '<div>Cover Page Footer</div>',
	},
	{
		displayName: 'Even Pages Header (HTML)',
		name: 'evenPagesHeaderHtml',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		default: '',
		description: 'HTML content for headers on even pages',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['html'],
			},
		},
	},
	{
		displayName: 'Even Pages Footer (HTML)',
		name: 'evenPagesFooterHtml',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		default: '',
		description: 'HTML content for footers on even pages',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['html'],
			},
		},
	},
	{
		displayName: 'Odd Pages Header (HTML)',
		name: 'oddPagesHeaderHtml',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		default: '',
		description: 'HTML content for headers on odd pages',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['html'],
			},
		},
	},
	{
		displayName: 'Odd Pages Footer (HTML)',
		name: 'oddPagesFooterHtml',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		default: '',
		description: 'HTML content for footers on odd pages',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
				contentType: ['html'],
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
				operation: [ActionConstants.UpdateHeadersFooters],
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
				operation: [ActionConstants.UpdateHeadersFooters],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output Binary Data Name',
		name: 'binaryDataName',
		type: 'string',
		default: 'data',
		description: 'Name for the binary data in the n8n output (used to access the processed files)',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.UpdateHeadersFooters],
			},
		},
	},
];

/**
 * Update Headers and Footers in Word documents using PDF4Me API
 * Process: Read Word file → Encode to base64 → Send API request → Poll for completion → Save updated Word file
 * Updates headers and footers in Word documents with configurable content for different page types
 */
export async function execute(this: IExecuteFunctions, index: number): Promise<INodeExecutionData[]> {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;
		const contentType = this.getNodeParameter('contentType', index, 'plain') as string;
		const cultureName = this.getNodeParameter('cultureName', index, 'en-US') as string;

		let docContent: string;
		let originalFileName = docName;

		// Handle different input data types
		if (inputDataType === 'binaryData') {
			// Get Word content from binary data
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
			// Use base64 content directly
			docContent = this.getNodeParameter('base64Content', index) as string;

			// Remove data URL prefix if present
			if (docContent.includes(',')) {
				docContent = docContent.split(',')[1];
			}
		} else if (inputDataType === 'url') {
			// Download Word file from URL
			const url = this.getNodeParameter('url', index) as string;

			if (!url || url.trim() === '') {
				throw new Error('URL is required when using URL input type');
			}

			try {
				// Download the file using n8n's helpers
				const response = await this.helpers.httpRequest({
					method: 'GET',
					url,
					encoding: 'arraybuffer',
					returnFullResponse: true,
				});

				// Convert to base64
				const buffer = Buffer.from(response.body as ArrayBuffer);
				docContent = buffer.toString('base64');

				// Try to extract filename from URL or Content-Disposition header
				const contentDisposition = response.headers['content-disposition'];
				if (contentDisposition) {
					const filenameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
					if (filenameMatch && filenameMatch[1]) {
						originalFileName = filenameMatch[1].replace(/['"]/g, '');
					}
				}

				// Fallback: extract filename from URL
				if (originalFileName === docName) {
					const urlParts = url.split('/');
					const urlFilename = urlParts[urlParts.length - 1].split('?')[0];
					if (urlFilename) {
						originalFileName = decodeURIComponent(urlFilename);
					}
				}
			} catch (error) {
				const errorMessage = error instanceof Error ? error.message : 'Unknown error';
				throw new Error(`Failed to download file from URL: ${errorMessage}`);
			}
		} else {
			throw new Error(`Unsupported input data type: ${inputDataType}`);
		}

		// Validate content
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

		// Build UpdateHeadersFootersAction based on content type
		const updateAction: IDataObject = {};

		if (contentType === 'plain') {
			// Plain text mode
			const allPagesHeader = this.getNodeParameter('allPagesHeader', index, '') as string;
			const allPagesFooter = this.getNodeParameter('allPagesFooter', index, '') as string;

			if (allPagesHeader) updateAction.AllPagesHeaderHtml = allPagesHeader;
			if (allPagesFooter) updateAction.AllPagesFooterHtml = allPagesFooter;
		} else if (contentType === 'html') {
			// HTML mode
			const allPagesHeaderHtml = this.getNodeParameter('allPagesHeaderHtml', index, '') as string;
			const allPagesFooterHtml = this.getNodeParameter('allPagesFooterHtml', index, '') as string;
			const firstPageHeaderHtml = this.getNodeParameter('firstPageHeaderHtml', index, '') as string;
			const firstPageFooterHtml = this.getNodeParameter('firstPageFooterHtml', index, '') as string;
			const evenPagesHeaderHtml = this.getNodeParameter('evenPagesHeaderHtml', index, '') as string;
			const evenPagesFooterHtml = this.getNodeParameter('evenPagesFooterHtml', index, '') as string;
			const oddPagesHeaderHtml = this.getNodeParameter('oddPagesHeaderHtml', index, '') as string;
			const oddPagesFooterHtml = this.getNodeParameter('oddPagesFooterHtml', index, '') as string;

			if (allPagesHeaderHtml) updateAction.AllPagesHeaderHtml = allPagesHeaderHtml;
			if (allPagesFooterHtml) updateAction.AllPagesFooterHtml = allPagesFooterHtml;
			if (firstPageHeaderHtml) updateAction.FirstPageHeaderHtml = firstPageHeaderHtml;
			if (firstPageFooterHtml) updateAction.FirstPageFooterHtml = firstPageFooterHtml;
			if (evenPagesHeaderHtml) updateAction.EvenPagesHeaderHtml = evenPagesHeaderHtml;
			if (evenPagesFooterHtml) updateAction.EvenPagesFooterHtml = evenPagesFooterHtml;
			if (oddPagesHeaderHtml) updateAction.OddPagesHeaderHtml = oddPagesHeaderHtml;
			if (oddPagesFooterHtml) updateAction.OddPagesFooterHtml = oddPagesFooterHtml;
		}

		// Add culture name if provided
		if (cultureName) {
			updateAction.CultureName = cultureName;
		}

		// Only add UpdateHeadersFootersAction if there's content to update
		if (Object.keys(updateAction).length > 0) {
			body.UpdateHeadersFootersAction = updateAction;
		} else {
			throw new Error('At least one header or footer content must be provided');
		}

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Word/UpdateHeadersFooters',
			body,
		);

		if (responseData) {
			// Generate filename for updated document
			const baseName = originalFileName ? originalFileName.replace(/\.[^.]*$/, '') : 'document';
			const fileName = `${baseName}_headers_footers_updated.docx`;

			// Handle the response - Word API returns JSON with embedded base64 file
			let wordBuffer: Buffer;

			// Check for Buffer first to properly narrow TypeScript types
			if (Buffer.isBuffer(responseData)) {
				// Direct binary response
				wordBuffer = responseData;
			} else if (typeof responseData === 'string') {
				// Base64 string response
				wordBuffer = Buffer.from(responseData, 'base64');
			} else if (typeof responseData === 'object' && responseData !== null) {
				// Try different possible response structures from IDataObject
				const response = responseData as IDataObject;

				// Check if the response has a document field
				if (response.document) {
					const document = response.document;
					// The document could be a string (base64) or an object with nested fields
					if (typeof document === 'string') {
						// Document itself is the base64 content
						wordBuffer = Buffer.from(document, 'base64');
					} else if (typeof document === 'object' && document !== null) {
						// Document is an object, extract base64 from possible fields
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
					// No document field, try other possible locations
					const docContent =
						(response.docContent as string) ||
						(response.docData as string) ||
						(response.content as string) ||
						(response.fileContent as string) ||
						(response.data as string);

					if (!docContent) {
						// If no known field found, log the structure for debugging
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

			const magicBytes = wordBuffer.toString('hex', 0, 4);
			if (magicBytes !== '504b0304') {
				throw new Error('Invalid DOCX file returned from API');
			}

			// Create binary data
			const binaryData = await this.helpers.prepareBinaryData(
				wordBuffer,
				fileName,
				'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			);

			// Determine the binary data name
			const binaryDataKey = binaryDataName || 'data';

			return [{
				json: {
					fileName,
					fileSize: wordBuffer.length,
					success: true,
					originalFileName,
					contentType,
					cultureName,
					message: 'Headers and footers updated successfully',
				},
				binary: {
					[binaryDataKey]: binaryData,
				},
			}];
		}

		throw new Error('No response data received from PDF4ME API');
	} catch (error) {
		// Re-throw the error with additional context
		const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
		throw new Error(`Update Headers and Footers failed: ${errorMessage}`);
	}
}

