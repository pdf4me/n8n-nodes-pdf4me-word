import type { IExecuteFunctions, IDataObject, INodeProperties } from 'n8n-workflow';
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
				operation: [ActionConstants.AddTextWatermarkToWord],
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
				operation: [ActionConstants.AddTextWatermarkToWord],
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
				operation: [ActionConstants.AddTextWatermarkToWord],
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
				operation: [ActionConstants.AddTextWatermarkToWord],
				inputDataType: ['url'],
			},
		},
	},
	// === WATERMARK SETTINGS ===
	{
		displayName: 'Watermark Text',
		name: 'watermarkText',
		type: 'string',
		required: true,
		default: 'CONFIDENTIAL',
		description: 'Text to appear as watermark (e.g., CONFIDENTIAL, DRAFT, INTERNAL USE ONLY)',
		placeholder: 'CONFIDENTIAL',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Orientation',
		name: 'orientation',
		type: 'options',
		required: true,
		default: 'Diagonal',
		description: 'Orientation of the watermark text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextWatermarkToWord],
			},
		},
		options: [
			{ name: 'Horizontal', value: 'Horizontal' },
			{ name: 'Vertical', value: 'Vertical' },
			{ name: 'Diagonal', value: 'Diagonal' },
			{ name: 'Upside-Down', value: 'Upside-Down' },
		],
	},
	{
		displayName: 'Font Family',
		name: 'fontFamily',
		type: 'options',
		default: 'Arial',
		description: 'Font family for the watermark text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextWatermarkToWord],
			},
		},
		options: [
			{ name: 'Arial', value: 'Arial' },
			{ name: 'Times New Roman', value: 'Times New Roman' },
			{ name: 'Courier New', value: 'Courier New' },
			{ name: 'Verdana', value: 'Verdana' },
			{ name: 'Calibri', value: 'Calibri' },
			{ name: 'Helvetica', value: 'Helvetica' },
			{ name: 'Georgia', value: 'Georgia' },
			{ name: 'Tahoma', value: 'Tahoma' },
		],
	},
	{
		displayName: 'Font Size',
		name: 'fontSize',
		type: 'number',
		default: 72,
		description: 'Font size for the watermark text',
		typeOptions: {
			minValue: 6,
			maxValue: 500,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Font Color',
		name: 'fontColor',
		type: 'color',
		default: '#808080',
		description: 'Color for the watermark text',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Semi Transparent',
		name: 'semiTransparent',
		type: 'boolean',
		default: true,
		description: 'Whether the watermark should be semi-transparent',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Rotation',
		name: 'rotation',
		type: 'number',
		default: 45,
		description: 'Rotation angle for the watermark in degrees',
		typeOptions: {
			minValue: -360,
			maxValue: 360,
		},
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextWatermarkToWord],
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
				operation: [ActionConstants.AddTextWatermarkToWord],
			},
		},
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'word_with_watermark.docx',
		description: 'Name for the processed Word file (will have watermark added)',
		placeholder: 'output.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextWatermarkToWord],
			},
		},
	},
	{
		displayName: 'Source Document Name',
		name: 'docName',
		type: 'string',
		default: 'myWordFile.docx',
		description: 'Name of the original Word file (for reference and processing)',
		placeholder: 'myWordFile.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.AddTextWatermarkToWord],
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
				operation: [ActionConstants.AddTextWatermarkToWord],
			},
		},
	},
];

/**
 * Add text watermark to Word documents using PDF4Me API
 * Process: Read Word file → Encode to base64 → Send API request → Poll for completion → Save Word file
 * Adds customizable text watermarks to Word documents with font, color, rotation, and orientation options
 */
export async function execute(this: IExecuteFunctions, index: number) {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;

		// Get watermark parameters
		const watermarkText = this.getNodeParameter('watermarkText', index) as string;
		const orientation = this.getNodeParameter('orientation', index, 'Diagonal') as string;
		const fontFamily = this.getNodeParameter('fontFamily', index, 'Arial') as string;
		const fontSize = this.getNodeParameter('fontSize', index, 72) as number;
		const fontColor = this.getNodeParameter('fontColor', index, '#808080') as string;
		const semiTransparent = this.getNodeParameter('semiTransparent', index, true) as boolean;
		const rotation = this.getNodeParameter('rotation', index, 45) as number;
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
			TextWatermarkAction: {
				WatermarkText: watermarkText,
				FontFamily: fontFamily,
				FontSize: fontSize,
				FontColor: fontColor,
				SemiTransparent: semiTransparent,
				Rotation: rotation,
				Orientation: orientation,
				CultureName: cultureName,
			},
			IsAsync: true,
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Word/WordAddTextWatermark',
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

			// The API returns JSON in format: { document: { docData: "base64..." }, ... }
			// or { docData: "base64..." } or similar structures
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
						originalFileName,
						watermarkText,
						orientation,
						fontFamily,
						fontSize,
						fontColor,
						semiTransparent,
						rotation,
						cultureName,
						message: 'Text watermark added to Word file successfully',
					},
					binary: {
						[binaryDataKey]: binaryData,
					},
				},
			];
		}

		throw new Error('No response data received from PDF4ME API');
	} catch (error) {
		// Re-throw the error with additional context
		const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
		throw new Error(`Add text watermark to Word failed: ${errorMessage}`);
	}
}

