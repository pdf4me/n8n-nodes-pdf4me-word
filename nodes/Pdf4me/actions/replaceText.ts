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
				operation: [ActionConstants.ReplaceText],
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
				operation: [ActionConstants.ReplaceText],
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
				operation: [ActionConstants.ReplaceText],
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
				operation: [ActionConstants.ReplaceText],
				inputDataType: ['url'],
			},
		},
	},
	// === TEXT REPLACEMENT PHRASES ===
	{
		displayName: 'Text Replacements',
		name: 'phrases',
		type: 'fixedCollection',
		typeOptions: {
			multipleValues: true,
		},
		required: true,
		default: {},
		description: 'Array of text replacements to perform',
		displayOptions: {
			show: {
				operation: [ActionConstants.ReplaceText],
			},
		},
		options: [
			{
				name: 'phrase',
				displayName: 'Replacement',
				values: [
					{
						displayName: 'Search Text',
						name: 'findText',
						type: 'string',
						required: true,
						default: '',
						description: 'Text to search for in the document',
						placeholder: 'Old Company',
					},
					{
						displayName: 'Replacement Text',
						name: 'replaceText',
						type: 'string',
						required: true,
						default: '',
						description: 'Text to replace the search text with',
						placeholder: 'New Company',
					},
					{
						displayName: 'Match Case',
						name: 'caseSensitive',
						type: 'boolean',
						default: false,
						description: 'Perform case-sensitive search',
					},
					{
						displayName: 'Match Whole Word',
						name: 'findWholeWordsOnly',
						type: 'boolean',
						default: false,
						description: 'Match whole words only',
					},
					{
						displayName: 'Use Regular Expressions',
						name: 'isExpression',
						type: 'boolean',
						default: false,
						description: 'Use regular expressions for search',
					},
					{
						displayName: 'Font Name',
						name: 'font',
						type: 'string',
						default: '',
						description: 'Font name for the replacement text (e.g., Arial, Times New Roman)',
						placeholder: 'Arial',
					},
					{
						displayName: 'Font Color',
						name: 'fontColor',
						type: 'color',
						default: '#000000',
						description: 'Text color for the replacement text',
					},
					{
						displayName: 'Font Size',
						name: 'fontSize',
						type: 'number',
						default: 0,
						description: 'Font size in points (0 = keep original size)',
						placeholder: '12',
					},
					{
						displayName: 'Background Color',
						name: 'backgroundColor',
						type: 'color',
						default: '#FFFFFF',
						description: 'Background color for the replacement text',
					},
					{
						displayName: 'Bold',
						name: 'bold',
						type: 'boolean',
						default: false,
						description: 'Apply bold formatting to replacement text',
					},
					{
						displayName: 'Italic',
						name: 'italic',
						type: 'boolean',
						default: false,
						description: 'Apply italic formatting to replacement text',
					},
					{
						displayName: 'Underline',
						name: 'underline',
						type: 'boolean',
						default: false,
						description: 'Apply underline formatting to replacement text',
					},
					{
						displayName: 'Strikethrough',
						name: 'strikethrough',
						type: 'boolean',
						default: false,
						description: 'Apply strikethrough formatting to replacement text',
					},
					{
						displayName: 'Double Strikethrough',
						name: 'doubleStrikethrough',
						type: 'boolean',
						default: false,
						description: 'Apply double strikethrough formatting to replacement text',
					},
					{
						displayName: 'Subscript',
						name: 'subscript',
						type: 'boolean',
						default: false,
						description: 'Apply subscript formatting to replacement text',
					},
					{
						displayName: 'Superscript',
						name: 'superscript',
						type: 'boolean',
						default: false,
						description: 'Apply superscript formatting to replacement text',
					},
					{
						displayName: 'Word Spacing',
						name: 'wordSpacing',
						type: 'number',
						default: 0,
						description: 'Word spacing in points (space between words, 0 = default)',
						placeholder: '0',
					},
				],
			},
		],
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
				operation: [ActionConstants.ReplaceText],
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
				operation: [ActionConstants.ReplaceText],
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
				operation: [ActionConstants.ReplaceText],
			},
		},
	},
];

/**
 * Replace Text in Word documents using PDF4Me API
 * Process: Read Word file → Encode to base64 → Send API request → Poll for completion → Save updated Word file
 * Replaces text in Word documents with configurable search options and formatting using a phrases array
 */
export async function execute(this: IExecuteFunctions, index: number): Promise<INodeExecutionData[]> {
	try {
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;
		const cultureName = this.getNodeParameter('cultureName', index, 'en-US') as string;

		// Get phrases array
		const phrasesData = this.getNodeParameter('phrases.phrase', index, []) as Array<{
			findText: string;
			replaceText: string;
			caseSensitive?: boolean;
			findWholeWordsOnly?: boolean;
			isExpression?: boolean;
			font?: string;
			fontColor?: string;
			fontSize?: number;
			backgroundColor?: string;
			bold?: boolean;
			italic?: boolean;
			underline?: boolean;
			strikethrough?: boolean;
			doubleStrikethrough?: boolean;
			subscript?: boolean;
			superscript?: boolean;
			wordSpacing?: number;
		}>;

		if (!phrasesData || phrasesData.length === 0) {
			throw new Error('At least one text replacement phrase is required');
		}

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

		// Build the phrases array with PascalCase field names for the API
		const phrases: IDataObject[] = phrasesData.map((phrase) => {
			const apiPhrase: IDataObject = {
				FindText: phrase.findText,
				ReplaceText: phrase.replaceText,
			};

			// Add optional search options
			if (phrase.caseSensitive !== undefined) {
				apiPhrase.CaseSensitive = phrase.caseSensitive;
			}
			if (phrase.findWholeWordsOnly !== undefined) {
				apiPhrase.FindWholeWordsOnly = phrase.findWholeWordsOnly;
			}
			if (phrase.isExpression !== undefined) {
				apiPhrase.IsExpression = phrase.isExpression;
			}

			// Add formatting options
			if (phrase.font && phrase.font.trim() !== '') {
				apiPhrase.Font = phrase.font;
			}
			if (phrase.fontColor && phrase.fontColor.trim() !== '') {
				apiPhrase.FontColor = phrase.fontColor;
			}
			if (phrase.fontSize !== undefined && phrase.fontSize > 0) {
				apiPhrase.FontSize = phrase.fontSize;
			}
			if (phrase.backgroundColor && phrase.backgroundColor.trim() !== '') {
				apiPhrase.BackgroundColor = phrase.backgroundColor;
			}
			if (phrase.bold !== undefined) {
				apiPhrase.Bold = phrase.bold;
			}
			if (phrase.italic !== undefined) {
				apiPhrase.Italic = phrase.italic;
			}
			if (phrase.underline !== undefined) {
				apiPhrase.Underline = phrase.underline;
			}
			if (phrase.strikethrough !== undefined) {
				apiPhrase.Strikethrough = phrase.strikethrough;
			}
			if (phrase.doubleStrikethrough !== undefined) {
				apiPhrase.DoubleStrikethrough = phrase.doubleStrikethrough;
			}
			if (phrase.subscript !== undefined) {
				apiPhrase.Subscript = phrase.subscript;
			}
			if (phrase.superscript !== undefined) {
				apiPhrase.Superscript = phrase.superscript;
			}
			if (phrase.wordSpacing !== undefined && phrase.wordSpacing !== 0) {
				apiPhrase.WordSpacing = phrase.wordSpacing;
			}

			return apiPhrase;
		});

		// Build the request body according to the API specification
		const body: IDataObject = {
			document: {
				Name: originalFileName,
			},
			docContent,
			Phrases: phrases,
		};

		// Add culture name if provided
		if (cultureName && cultureName.trim() !== '') {
			body.CultureName = cultureName;
		}

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Word/ReplaceText',
			body,
		);

		if (responseData) {
			// Generate filename for updated document
			const baseName = originalFileName ? originalFileName.replace(/\.[^.]*$/, '') : 'document';
			const fileName = `${baseName}_text_replaced.docx`;

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
					phrasesCount: phrases.length,
					cultureName,
					message: 'Text replaced successfully',
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
		throw new Error(`Replace Text failed: ${errorMessage}`);
	}
}
