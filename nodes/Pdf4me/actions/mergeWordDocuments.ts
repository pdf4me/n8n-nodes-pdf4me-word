import type { IExecuteFunctions, IDataObject, INodeExecutionData, INodeProperties } from 'n8n-workflow';
import {
	pdf4meAsyncRequest,
	ActionConstants,
} from '../GenericFunctions';

export const description: INodeProperties[] = [
	// === DOCUMENTS SOURCE ===
	{
		displayName: 'Documents Source',
		name: 'docsSource',
		type: 'options',
		default: 'items',
		description: 'How to provide documents to merge',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
			},
		},
		options: [
			{
				name: 'From Previous Node Items',
				value: 'items',
				description: 'Merge documents from multiple items in previous node (each item should have a binary file)',
			},
			{
				name: 'Manual Entry',
				value: 'parameter',
				description: 'Manually specify documents',
			},
		],
	},
	{
		displayName: 'Documents',
		name: 'documents',
		type: 'fixedCollection',
		typeOptions: {
			multipleValues: true,
		},
		default: {},
		description: 'Word documents to merge',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
				docsSource: ['parameter'],
			},
		},
		options: [
			{
				name: 'document',
				displayName: 'Document',
				values: [
					{
						displayName: 'Input Method',
						name: 'inputMethod',
						type: 'options',
						default: 'base64',
						options: [
							{
								name: 'Base64 Encoded',
								value: 'base64',
							},
							{
								name: 'Download from URL',
								value: 'url',
							},
						],
					},
					{
						displayName: 'Document Name',
						name: 'fileName',
						type: 'string',
						default: '',
						description: 'Name of the document file',
						placeholder: 'document2.docx',
					},
					{
						displayName: 'Base64 Content',
						name: 'base64Content',
						type: 'string',
						typeOptions: {
							alwaysOpenEditWindow: true,
						},
						default: '',
						description: 'Base64 encoded string containing the Word file data',
						displayOptions: {
							show: {
								inputMethod: ['base64'],
							},
						},
					},
					{
						displayName: 'URL',
						name: 'url',
						type: 'string',
						default: '',
						description: 'URL to download the Word file from',
						displayOptions: {
							show: {
								inputMethod: ['url'],
							},
						},
					},
					{
						displayName: 'Sort Position',
						name: 'sortPosition',
						type: 'number',
						default: 0,
						description: 'Position in merge order (0 = first, 1 = second, etc.)',
					},
				],
			},
		],
	},
	{
		displayName: 'Binary Property Name',
		name: 'binaryPropertyName',
		type: 'string',
		default: 'data',
		description: 'Name of the binary property containing Word files (for items mode)',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
				docsSource: ['items'],
			},
		},
	},
	// === MERGE OPTIONS ===
	{
		displayName: 'Format Mode',
		name: 'formatMode',
		type: 'options',
		default: 'KeepSourceFormatting',
		description: 'Select how to handle merging Microsoft Word document styles',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
			},
		},
		options: [
			{
				name: 'Keep Source Formatting',
				value: 'KeepSourceFormatting',
				description: 'Maintain all formatting from source documents',
			},
			{
				name: 'Keep Different Styles',
				value: 'KeepDifferentStyles',
				description: 'Keep only different styles from each document',
			},
			{
				name: 'Use Destination Styles',
				value: 'UseDestinationStyles',
				description: 'Apply destination document styles to all content',
			},
		],
	},
	{
		displayName: 'Compliance Level',
		name: 'complianceLevel',
		type: 'options',
		default: 'Transitional',
		description: 'Document compliance level',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
			},
		},
		options: [
			{
				name: 'ECMA',
				value: 'ECMA',
				description: 'ECMA compliance level',
			},
			{
				name: 'Transitional',
				value: 'Transitional',
				description: 'Transitional compliance level (recommended)',
			},
			{
				name: 'Strict',
				value: 'Strict',
				description: 'Strict compliance level',
			},
			{
				name: 'Custom',
				value: 'Custom',
				description: 'Custom compliance level',
			},
		],
	},
	// === OUTPUT SETTINGS ===
	{
		displayName: 'Output File Name',
		name: 'outputFileName',
		type: 'string',
		default: 'merged_document.docx',
		description: 'Name for the merged Word file',
		placeholder: 'merged.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
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
				operation: [ActionConstants.MergeWordDocuments],
			},
		},
	},
];

/**
 * Helper function to get document content from different input types
 */
async function getDocumentContentFromInput(
	this: IExecuteFunctions,
	index: number,
	inputMethod: string,
	base64Content: string | undefined,
	url: string | undefined,
	binaryPropertyName: string | undefined,
	defaultFileName: string,
): Promise<{ content: string; fileName: string }> {
	let docContent: string;
	let fileName = defaultFileName;

	if (inputMethod === 'binaryData' && binaryPropertyName) {
		const item = this.getInputData(index);
		if (!item[0]?.binary || !item[0].binary[binaryPropertyName]) {
			throw new Error(`No binary data found in property '${binaryPropertyName}'`);
		}
		const binaryData = item[0].binary[binaryPropertyName];
		const buffer = await this.helpers.getBinaryDataBuffer(index, binaryPropertyName);
		docContent = buffer.toString('base64');
		if (binaryData.fileName) {
			fileName = binaryData.fileName;
		}
	} else if (inputMethod === 'base64' && base64Content) {
		docContent = base64Content;
		if (docContent.includes(',')) {
			docContent = docContent.split(',')[1];
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
		docContent = buffer.toString('base64');
		const contentDisposition = response.headers['content-disposition'];
		if (contentDisposition) {
			const filenameMatch = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
			if (filenameMatch && filenameMatch[1]) {
				fileName = filenameMatch[1].replace(/['"]/g, '');
			}
		}
		if (fileName === defaultFileName) {
			const urlParts = url.split('/');
			const urlFilename = urlParts[urlParts.length - 1].split('?')[0];
			if (urlFilename) {
				fileName = decodeURIComponent(urlFilename);
			}
		}
	} else {
		throw new Error(`Invalid input method or missing content: ${inputMethod}`);
	}

	if (!docContent || docContent.trim() === '') {
		throw new Error('Document content is required');
	}

	return { content: docContent, fileName };
}

/**
 * Merge multiple Word documents using PDF4Me API
 * Process: Read all Word files → Encode to base64 → Send API request → Poll for completion → Save merged Word file
 * Merges multiple Word documents into a single document with configurable options
 */
export async function execute(this: IExecuteFunctions, index: number): Promise<INodeExecutionData[]> {
	try {
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const docsSource = this.getNodeParameter('docsSource', index, 'items') as string;
		const formatMode = this.getNodeParameter('formatMode', index, 'KeepSourceFormatting') as string;
		const complianceLevel = this.getNodeParameter('complianceLevel', index, 'Transitional') as string;

		// Collect all documents to merge
		const documentsToMerge: Array<{ FileContent: string; FileName: string; SortPosition: number }> = [];

		// Get documents based on source
		if (docsSource === 'items') {
			// Get documents from multiple input items
			const binaryPropertyName = this.getNodeParameter('binaryPropertyName', index, 'data') as string;
			const items = this.getInputData();

			for (let i = 0; i < items.length; i++) {
				const item = items[i];
				if (!item?.binary || !item.binary[binaryPropertyName]) continue;

				try {
					const binaryData = item.binary[binaryPropertyName];
					const buffer = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
					const docContent = buffer.toString('base64');
					const fileName = binaryData.fileName || `document${i + 1}.docx`;

					documentsToMerge.push({
						FileContent: docContent,
						FileName: fileName,
						SortPosition: i,
					});
				} catch (error) {
					// Skip items that fail to load
					continue;
				}
			}
		} else {
			// Get documents from parameter
			const documents = this.getNodeParameter('documents.document', index, []) as Array<{
				inputMethod: string;
				fileName: string;
				base64Content?: string;
				url?: string;
				sortPosition: number;
			}>;

			for (const docConfig of documents) {
				try {
					const doc = await getDocumentContentFromInput.call(
						this,
						index,
						docConfig.inputMethod,
						docConfig.base64Content,
						docConfig.url,
						undefined,
						docConfig.fileName || 'document.docx',
					);

					documentsToMerge.push({
						FileContent: doc.content,
						FileName: doc.fileName,
						SortPosition: docConfig.sortPosition || documentsToMerge.length,
					});
				} catch (error) {
					// Skip documents that fail to load
					continue;
				}
			}
		}

		// Sort by SortPosition to ensure correct merge order
		documentsToMerge.sort((a, b) => a.SortPosition - b.SortPosition);

		if (documentsToMerge.length < 2) {
			throw new Error('At least 2 documents are required for merging');
		}

		// Build the request body according to the API specification
		const Documents = documentsToMerge.map(doc => ({
			Filename: doc.FileName,
			DocContent: doc.FileContent,
			SortPosition: doc.SortPosition,
			FormatMode: formatMode,
		}));

		const body: IDataObject = {
			Documents,
			MergeOptions: {
				ComplianceLevel: complianceLevel,
			},
		};

		// Send the request to the API
		const responseData = await pdf4meAsyncRequest.call(
			this,
			'/office/ApiV2Word/MergeDocuments',
			body,
		);

		if (responseData) {
			// Generate filename if not provided
			let fileName = outputFileName;
			if (!fileName || fileName.trim() === '') {
				fileName = 'merged_document.docx';
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

			// Validate Word file format
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
						documentsMerged: documentsToMerge.length,
						documentNames: documentsToMerge.map(d => d.FileName),
						message: `Successfully merged ${documentsToMerge.length} documents`,
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
		throw new Error(`Merge Word documents failed: ${errorMessage}`);
	}
}

