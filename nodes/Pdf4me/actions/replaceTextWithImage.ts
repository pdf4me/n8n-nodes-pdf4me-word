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
				operation: [ActionConstants.ReplaceTextWithImage],
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
				operation: [ActionConstants.ReplaceTextWithImage],
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
				operation: [ActionConstants.ReplaceTextWithImage],
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
				operation: [ActionConstants.ReplaceTextWithImage],
				inputDataType: ['url'],
			},
		},
	},
	// === REPLACEMENT IMAGE SETTINGS ===
	{
		displayName: 'Image Input Method',
		name: 'imageInputDataType',
		type: 'options',
		required: true,
		default: 'binaryData',
		description: 'Choose how to provide the replacement image',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
			},
		},
		options: [
			{
				name: 'From Previous Node (Binary Data)',
				value: 'binaryData',
				description: 'Use image file passed from a previous n8n node',
			},
			{
				name: 'Base64 Encoded String',
				value: 'base64',
				description: 'Provide image content as base64 encoded string',
			},
			{
				name: 'Download from URL',
				value: 'url',
				description: 'Download image directly from a web URL',
			},
		],
	},
	{
		displayName: 'Image Binary Data Property Name',
		name: 'imageBinaryPropertyName',
		type: 'string',
		required: true,
		default: 'data',
		description: 'Name of the binary property containing the replacement image',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
				imageInputDataType: ['binaryData'],
			},
		},
	},
	{
		displayName: 'Base64 Encoded Image Content',
		name: 'imageBase64Content',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		required: true,
		default: '',
		description: 'Base64 encoded string containing the image data (PNG, JPG, etc.)',
		placeholder: 'iVBORw0KGgoAAAANS...',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
				imageInputDataType: ['base64'],
			},
		},
	},
	{
		displayName: 'Image URL',
		name: 'imageUrl',
		type: 'string',
		required: true,
		default: '',
		description: 'URL to download the replacement image from',
		placeholder: 'https://example.com/image.png',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
				imageInputDataType: ['url'],
			},
		},
	},
	// === SEARCH AND REPLACEMENT SETTINGS ===
	{
		displayName: 'Search Text',
		name: 'findText',
		type: 'string',
		default: '',
		description: 'Text to search for and replace with image (leave empty to replace all text)',
		placeholder: 'Old Company Name',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
			},
		},
	},
	// === IMAGE SIZE SETTINGS ===
	{
		displayName: 'Width (Points)',
		name: 'width',
		type: 'number',
		default: '',
		description: 'Image width in points (optional, maintains aspect ratio if only one dimension provided)',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
			},
		},
	},
	{
		displayName: 'Height (Points)',
		name: 'height',
		type: 'number',
		default: '',
		description: 'Image height in points (optional, maintains aspect ratio if only one dimension provided)',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
			},
		},
	},
	{
		displayName: 'Maintain Aspect Ratio',
		name: 'maintainAspectRatio',
		type: 'boolean',
		default: true,
		description: 'Maintain image aspect ratio when resizing',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
			},
		},
	},
	// === PAGE FILTERING SETTINGS ===
	{
		displayName: 'Skip First Page',
		name: 'skipFirstPage',
		type: 'boolean',
		default: false,
		description: 'Skip replacement on the first page',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
			},
		},
	},
	{
		displayName: 'Apply To',
		name: 'applyTo',
		type: 'options',
		default: 'all',
		description: 'Apply replacement to specific page types',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
			},
		},
		options: [
			{
				name: 'All Pages',
				value: 'all',
			},
			{
				name: 'Odd Pages',
				value: 'odd',
			},
			{
				name: 'Even Pages',
				value: 'even',
			},
			{
				name: 'First Page',
				value: 'first',
			},
			{
				name: 'Last Page',
				value: 'last',
			},
		],
	},
	{
		displayName: 'Page Numbers',
		name: 'pageNumbers',
		type: 'string',
		default: '',
		description: 'Comma-separated page numbers or ranges (e.g., "1,3,5-7")',
		placeholder: '1,3,5-7',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
			},
		},
	},
	{
		displayName: 'Ignore Page Numbers',
		name: 'ignorePageNumbers',
		type: 'string',
		default: '',
		description: 'Comma-separated page numbers to skip (e.g., "2,4")',
		placeholder: '2,4',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceTextWithImage],
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
				operation: [ActionConstants.ReplaceTextWithImage],
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
				operation: [ActionConstants.ReplaceTextWithImage],
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
				operation: [ActionConstants.ReplaceTextWithImage],
			},
		},
	},
];

/**
 * Helper function to get image content from different input types
 */
async function getImageContentFromInput(
	this: IExecuteFunctions,
	index: number,
	inputMethod: string,
	base64Content: string | undefined,
	url: string | undefined,
	binaryPropertyName: string | undefined,
): Promise<string> {
	let imageContent: string;

	if (inputMethod === 'binaryData' && binaryPropertyName) {
		const item = this.getInputData(index);
		if (!item[0]?.binary || !item[0].binary[binaryPropertyName]) {
			throw new Error(`No binary image data found in property '${binaryPropertyName}'`);
		}
		const buffer = await this.helpers.getBinaryDataBuffer(index, binaryPropertyName);
		imageContent = buffer.toString('base64');
	} else if (inputMethod === 'base64' && base64Content) {
		imageContent = base64Content;
		if (imageContent.includes(',')) {
			imageContent = imageContent.split(',')[1];
		}
	} else if (inputMethod === 'url' && url) {
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
		imageContent = buffer.toString('base64');
	} else {
		throw new Error(`Invalid image input method or missing content: ${inputMethod}`);
	}

	if (!imageContent || imageContent.trim() === '') {
		throw new Error('Image content is required');
	}

	return imageContent;
}

/**
 * Replace Text with Image in Word documents using PDF4Me API
 * Process: Read Word file & image → Encode to base64 → Send API request → Poll for completion → Save updated Word file
 * Replaces specified text in Word documents with images, with configurable size, aspect ratio, and page filtering options
 */
export async function execute(this: IExecuteFunctions, index: number): Promise<INodeExecutionData[]> {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;
		const findText = this.getNodeParameter('findText', index, '') as string;
		const cultureName = this.getNodeParameter('cultureName', index, 'en-US') as string;

		// Get image input method
		const imageInputDataType = this.getNodeParameter('imageInputDataType', index) as string;

		// Get image size parameters
		const width = this.getNodeParameter('width', index, '') as string | number;
		const height = this.getNodeParameter('height', index, '') as string | number;
		const maintainAspectRatio = this.getNodeParameter('maintainAspectRatio', index, true) as boolean;

		// Get page filtering parameters
		const skipFirstPage = this.getNodeParameter('skipFirstPage', index, false) as boolean;
		const applyTo = this.getNodeParameter('applyTo', index, 'all') as string;
		const pageNumbers = this.getNodeParameter('pageNumbers', index, '') as string;
		const ignorePageNumbers = this.getNodeParameter('ignorePageNumbers', index, '') as string;

		let docContent: string;
		let originalFileName = docName;

		// Handle different Word document input data types
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
			} catch (error) {
				const errorMessage = error instanceof Error ? error.message : 'Unknown error';
				throw new Error(`Failed to download file from URL: ${errorMessage}`);
			}
		} else {
			throw new Error(`Unsupported input data type: ${inputDataType}`);
		}

		// Validate document content
		if (!docContent || docContent.trim() === '') {
			throw new Error('Word content is required');
		}

		// Get replacement image content
		const imageContent = await getImageContentFromInput.call(
			this,
			index,
			imageInputDataType,
			imageInputDataType === 'base64' ? (this.getNodeParameter('imageBase64Content', index) as string) : undefined,
			imageInputDataType === 'url' ? (this.getNodeParameter('imageUrl', index) as string) : undefined,
			imageInputDataType === 'binaryData' ? (this.getNodeParameter('imageBinaryPropertyName', index) as string) : undefined,
		);

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				Name: originalFileName,
			},
			docContent,
			ImageContent: imageContent,
			MaintainAspectRatio: maintainAspectRatio,
			SkipFirstPage: skipFirstPage,
			ApplyTo: applyTo,
			CultureName: cultureName,
		};

		// Add FindText if provided
		if (findText && findText.trim() !== '') {
			body.FindText = findText;
		}

		// Add width and height if provided
		if (width && width !== '' && !isNaN(Number(width))) {
			body.Width = Number(width);
		}
		if (height && height !== '' && !isNaN(Number(height))) {
			body.Height = Number(height);
		}

		// Add page filtering options if provided
		if (pageNumbers && pageNumbers.trim() !== '') {
			body.PageNumbers = pageNumbers.trim();
		}
		if (ignorePageNumbers && ignorePageNumbers.trim() !== '') {
			body.IgnorePageNumbers = ignorePageNumbers.trim();
		}

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Word/ReplaceTextWithImage',
			body,
		);

		if (responseData) {
			// Generate filename for updated document
			const baseName = originalFileName ? originalFileName.replace(/\.[^.]*$/, '') : 'document';
			const fileName = `${baseName}_text_replaced_with_image.docx`;

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
					findText: findText || undefined,
					width: width || undefined,
					height: height || undefined,
					maintainAspectRatio,
					skipFirstPage,
					applyTo,
					pageNumbers: pageNumbers || undefined,
					ignorePageNumbers: ignorePageNumbers || undefined,
					cultureName,
					message: 'Text replaced with image successfully',
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
		throw new Error(`Replace Text with Image failed: ${errorMessage}`);
	}
}

