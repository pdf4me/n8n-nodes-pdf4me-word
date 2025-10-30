import type { IExecuteFunctions, IDataObject, INodeProperties, INodeExecutionData } from 'n8n-workflow';
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
				operation: [ActionConstants.AddImageWatermarkToWord],
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
				operation: [ActionConstants.AddImageWatermarkToWord],
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
				operation: [ActionConstants.AddImageWatermarkToWord],
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
				operation: [ActionConstants.AddImageWatermarkToWord],
				inputDataType: ['url'],
			},
		},
	},
	{
		displayName: 'Document Name',
		name: 'docName',
		type: 'string',
		default: 'document.docx',
		description: 'Name of the Word file',
		placeholder: 'document.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
			},
		},
	},
	// === WATERMARK IMAGE SETTINGS ===
	{
		displayName: 'Watermark Image Input Method',
		name: 'imageInputDataType',
		type: 'options',
		required: true,
		default: 'binaryData',
		description: 'Choose how to provide the watermark image',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
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
		description: 'Name of the binary property containing the watermark image',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
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
				operation: [ActionConstants.AddImageWatermarkToWord],
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
		description: 'URL to download the watermark image from',
		placeholder: 'https://example.com/watermark.png',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
				imageInputDataType: ['url'],
			},
		},
	},
	// === WATERMARK OPTIONS ===
	{
		displayName: 'Scale',
		name: 'scale',
		type: 'number',
		default: 1.0,
		description: 'Scaling factor for watermark (0.1 to 10, e.g., 0.5 for 50%, 2.0 for 200%)',
		typeOptions: {
			minValue: 0.1,
			maxValue: 10,
			numberStepSize: 0.1,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Width (Points)',
		name: 'width',
		type: 'number',
		default: '',
		description: 'Watermark width in points (overrides Scale if provided). Leave empty to use Scale.',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Height (Points)',
		name: 'height',
		type: 'number',
		default: '',
		description: 'Watermark height in points (overrides Scale if provided). Leave empty to use Scale.',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Center Alignment',
		name: 'alignImage',
		type: 'boolean',
		default: true,
		description: 'If true, center image on page; if false, align top-left',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Semi-Transparent',
		name: 'semiTransparent',
		type: 'boolean',
		default: false,
		description: 'If true, watermark opacity is 50%; if false, 100%',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Culture Name',
		name: 'cultureName',
		type: 'string',
		default: 'en-US',
		description: 'Culture/locale name for document metadata (e.g., en-US, fr-FR, de-DE)',
		placeholder: 'en-US',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'word_with_watermark.docx',
		description: 'Name for the Word file with watermark',
		placeholder: 'watermarked.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddImageWatermarkToWord],
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
				operation: [ActionConstants.AddImageWatermarkToWord],
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
 * Add image watermark to Word documents using PDF4Me API
 * Process: Read Word file & image → Encode to base64 → Send API request → Poll for completion → Save Word file
 * Adds image watermarks to Word documents with configurable scale, size, alignment, and transparency
 */
export async function execute(this: IExecuteFunctions, index: number): Promise<INodeExecutionData[]> {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;

		// Get watermark image input method
		const imageInputDataType = this.getNodeParameter('imageInputDataType', index) as string;

		// Get watermark parameters
		const scale = this.getNodeParameter('scale', index, 1.0) as number;
		const width = this.getNodeParameter('width', index, '') as string | number;
		const height = this.getNodeParameter('height', index, '') as string | number;
		const alignImage = this.getNodeParameter('alignImage', index, true) as boolean;
		const semiTransparent = this.getNodeParameter('semiTransparent', index, false) as boolean;
		const cultureName = this.getNodeParameter('cultureName', index, 'en-US') as string;

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

		// Get watermark image content
		const watermarkFileContent = await getImageContentFromInput.call(
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
				name: originalFileName,
			},
			docContent,
			AddWatermarkAction: {
				WatermarkFileContent: watermarkFileContent,
				Scale: scale,
				AlignImage: alignImage,
				SemiTransparent: semiTransparent,
				CultureName: cultureName,
			},
		};

		// Add width and height if provided (override scale)
		if (width && width !== '' && !isNaN(Number(width))) {
			(body.AddWatermarkAction as IDataObject).Width = Number(width);
		}
		if (height && height !== '' && !isNaN(Number(height))) {
			(body.AddWatermarkAction as IDataObject).Height = Number(height);
		}

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Word/AddImageWatermark',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				const baseName = originalFileName
					? originalFileName.replace(/\.[^.]*$/, '')
					: 'word_with_watermark';
				fileName = `${baseName}.docx`;
			}

			// Ensure .docx extension
			if (!fileName.toLowerCase().endsWith('.docx')) {
				fileName = `${fileName.replace(/\.[^.]*$/, '')}.docx`;
			}

			// Handle the response - Word API returns JSON with embedded base64 file
			let wordBuffer: Buffer;

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

			const binaryDataKey = binaryDataName || 'data';

			return [
				{
					json: {
						fileName,
						fileSize: wordBuffer.length,
						success: true,
						originalFileName,
						scale,
						width: width || undefined,
						height: height || undefined,
						alignImage,
						semiTransparent,
						cultureName,
						message: 'Word document with image watermark created successfully',
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
		throw new Error(`Add image watermark to Word document failed: ${errorMessage}`);
	}
}

