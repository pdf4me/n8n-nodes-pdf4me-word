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
				operation: [ActionConstants.SplitWordDocument],
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
				operation: [ActionConstants.SplitWordDocument],
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
				operation: [ActionConstants.SplitWordDocument],
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
				operation: [ActionConstants.SplitWordDocument],
				inputDataType: ['url'],
			},
		},
	},
	// === SPLIT SETTINGS ===
	{
		displayName: 'Split Type',
		name: 'splitType',
		type: 'options',
		required: true,
		default: 'Pages',
		description: 'Type of split operation to perform',
		displayOptions: {
			show: {
				operation: [ActionConstants.SplitWordDocument],
			},
		},
		options: [
			{
				name: 'Pages',
				value: 'Pages',
				description: 'Split by page ranges (e.g., "1-3,5,7-9")',
			},
			{
				name: 'Sections',
				value: 'Sections',
				description: 'Split by document sections',
			},
			{
				name: 'Headings',
				value: 'Headings',
				description: 'Split by heading levels',
			},
			{
				name: 'Custom',
				value: 'Custom',
				description: 'Custom split configuration',
			},
		],
	},
	{
		displayName: 'Page Ranges',
		name: 'pageRanges',
		type: 'string',
		default: '',
		description: 'Comma-separated page ranges (e.g., "1-3,5,7-9" or "1,2,3"). Used when Split Type is Pages or Custom',
		placeholder: '1-3,5,7-9',
		displayOptions: {
			show: {
				operation: [ActionConstants.SplitWordDocument],
				splitType: ['Pages', 'Custom'],
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
				operation: [ActionConstants.SplitWordDocument],
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
				operation: [ActionConstants.SplitWordDocument],
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
				operation: [ActionConstants.SplitWordDocument],
			},
		},
	},
];

/**
 * Split Word documents using PDF4Me API
 * Process: Read Word file → Encode to base64 → Send API request → Poll for completion → Save split Word files
 * Splits Word documents by pages, sections, headings, or custom ranges into multiple separate documents
 */
export async function execute(this: IExecuteFunctions, index: number): Promise<INodeExecutionData[]> {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;
		const splitType = this.getNodeParameter('splitType', index) as string;
		const pageRanges = this.getNodeParameter('pageRanges', index, '') as string;
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
				Name: originalFileName,
			},
			docContent,
			splitType,
			cultureName,
		};

		// Add PageRanges if provided and SplitType is Pages or Custom
		if (pageRanges && pageRanges.trim() !== '' && (splitType === 'Pages' || splitType === 'Custom')) {
			body.pageRanges = pageRanges.trim();
		}

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Word/SplitDocument',
			body,
		);

		if (responseData) {
			// The API returns JSON with documents array
			const response = responseData as IDataObject;

			// Check if operation was successful
			const success = response.Success as boolean;
			if (!success) {
				const errorMessage = (response.ErrorMessage as string) || 'Unknown error';
				const errors = (response.Errors as string[]) || [];
				throw new Error(`Split operation failed: ${errorMessage}${errors.length > 0 ? '. Errors: ' + errors.join(', ') : ''}`);
			}

			// Get the array of split documents
			const documents = response.documents as string[];
			if (!documents || !Array.isArray(documents) || documents.length === 0) {
				throw new Error('No documents returned from split operation');
			}

			// Process each split document and return as separate items
			const results: INodeExecutionData[] = [];
			const baseName = originalFileName ? originalFileName.replace(/\.[^.]*$/, '') : 'document';

			for (let i = 0; i < documents.length; i++) {
				const docBase64 = documents[i];
				if (!docBase64 || typeof docBase64 !== 'string') {
					continue;
				}

				// Convert base64 to buffer
				const wordBuffer = Buffer.from(docBase64, 'base64');

				// Validate Word file format
				if (wordBuffer.length < 1000) {
					continue; // Skip invalid files
				}

				const magicBytes = wordBuffer.toString('hex', 0, 4);
				if (magicBytes !== '504b0304') {
					continue; // Skip invalid DOCX files
				}

				// Generate filename for split document
				const fileName = `${baseName}_part${i + 1}.docx`;

				// Create binary data
				const binaryData = await this.helpers.prepareBinaryData(
					wordBuffer,
					fileName,
					'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				);

				// Determine the binary data name
				const binaryDataKey = binaryDataName || 'data';

				results.push({
					json: {
						fileName,
						fileSize: wordBuffer.length,
						partNumber: i + 1,
						totalParts: documents.length,
						success: true,
						originalFileName,
						splitType,
						pageRanges: pageRanges || undefined,
						cultureName,
						message: `Split document part ${i + 1} of ${documents.length}`,
					},
					binary: {
						[binaryDataKey]: binaryData,
					},
				});
			}

			if (results.length === 0) {
				throw new Error('No valid documents could be extracted from split operation');
			}

			return results;
		}

		throw new Error('No response data received from PDF4ME API');
	} catch (error) {
		// Re-throw the error with additional context
		const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
		throw new Error(`Split Word document failed: ${errorMessage}`);
	}
}

