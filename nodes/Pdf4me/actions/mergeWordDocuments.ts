import type { IExecuteFunctions, IDataObject, INodeExecutionData, INodeProperties } from 'n8n-workflow';
import {
	pdf4meAsyncRequest,
	ActionConstants,
} from '../GenericFunctions';

export const description: INodeProperties[] = [
	// === MAIN DOCUMENT INPUT SETTINGS ===
	{
		displayName: 'Main Document Input Method',
		name: 'inputDataType',
		type: 'options',
		required: true,
		default: 'binaryData',
		description: 'Choose how to provide the main Word file (first document)',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
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
		displayName: 'Main Document Binary Property Name',
		name: 'binaryPropertyName',
		type: 'string',
		required: true,
		default: 'data',
		description: 'Name of the binary property containing the main Word file',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
				inputDataType: ['binaryData'],
			},
		},
	},
	{
		displayName: 'Main Document Base64 Content',
		name: 'base64Content',
		type: 'string',
		typeOptions: {
			alwaysOpenEditWindow: true,
		},
		required: true,
		default: '',
		description: 'Base64 encoded string containing the main Word file data',
		placeholder: 'UEsDBAoAAAAAAIdO4kAAAAAAAAAAAAAAAAAJAAAAZG9jUHJvcHMv...',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
				inputDataType: ['base64'],
			},
		},
	},
	{
		displayName: 'Main Document URL',
		name: 'url',
		type: 'string',
		required: true,
		default: '',
		description: 'URL to download the main Word file from',
		placeholder: 'https://example.com/document1.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
				inputDataType: ['url'],
			},
		},
	},
	{
		displayName: 'Main Document Name',
		name: 'docName',
		type: 'string',
		default: 'document1.docx',
		description: 'Name of the main Word file',
		placeholder: 'document1.docx',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
			},
		},
	},
	// === ADDITIONAL DOCUMENTS TO MERGE ===
	{
		displayName: 'Additional Documents Source',
		name: 'additionalDocsSource',
		type: 'options',
		default: 'parameter',
		description: 'How to provide additional documents to merge',
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
				description: 'Manually specify additional documents',
			},
		],
	},
	{
		displayName: 'Additional Documents',
		name: 'additionalDocuments',
		type: 'fixedCollection',
		typeOptions: {
			multipleValues: true,
		},
		default: {},
		description: 'Additional Word documents to merge with the main document',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
				additionalDocsSource: ['parameter'],
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
						default: 1,
						description: 'Position in merge order (0 = first/main document, 1 = second, etc.)',
					},
				],
			},
		],
	},
	{
		displayName: 'Additional Documents Binary Property Name',
		name: 'additionalBinaryPropertyName',
		type: 'string',
		default: 'data',
		description: 'Name of the binary property containing additional Word files (for items mode)',
		placeholder: 'data',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
				additionalDocsSource: ['items'],
			},
		},
	},
	// === MERGE OPTIONS ===
	{
		displayName: 'Maintain Formatting',
		name: 'maintainFormatting',
		type: 'boolean',
		default: true,
		description: 'Whether to maintain formatting from source documents',
		displayOptions: {
			show: {
				operation: [ActionConstants.MergeWordDocuments],
			},
		},
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
		const inputDataType = this.getNodeParameter('inputDataType', index) as string;
		const docName = this.getNodeParameter('docName', index) as string;
		const binaryDataName = this.getNodeParameter('binaryDataName', index) as string;
		const outputFileName = this.getNodeParameter('outputFileName', index) as string;
		const maintainFormatting = this.getNodeParameter('maintainFormatting', index, true) as boolean;
		const additionalDocsSource = this.getNodeParameter('additionalDocsSource', index, 'parameter') as string;

		// Get main document
		const mainDoc = await getDocumentContentFromInput.call(
			this,
			index,
			inputDataType,
			inputDataType === 'base64' ? (this.getNodeParameter('base64Content', index) as string) : undefined,
			inputDataType === 'url' ? (this.getNodeParameter('url', index) as string) : undefined,
			inputDataType === 'binaryData' ? (this.getNodeParameter('binaryPropertyName', index) as string) : undefined,
			docName,
		);

		// Collect all documents to merge
		const documentsToMerge: Array<{ FileContent: string; FileName: string; SortPosition: number }> = [];

		// Add main document first (position 0)
		documentsToMerge.push({
			FileContent: mainDoc.content,
			FileName: mainDoc.fileName,
			SortPosition: 0,
		});

		// Get additional documents
		if (additionalDocsSource === 'items') {
			// Get documents from multiple input items
			const additionalBinaryPropertyName = this.getNodeParameter('additionalBinaryPropertyName', index, 'data') as string;
			const items = this.getInputData();

			for (let i = 0; i < items.length; i++) {
				if (i === index) continue; // Skip the main document item
				const item = items[i];
				if (!item?.binary || !item.binary[additionalBinaryPropertyName]) continue;

				try {
					const binaryData = item.binary[additionalBinaryPropertyName];
					const buffer = await this.helpers.getBinaryDataBuffer(i, additionalBinaryPropertyName);
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
			const additionalDocuments = this.getNodeParameter('additionalDocuments.document', index, []) as Array<{
				inputMethod: string;
				fileName: string;
				base64Content?: string;
				url?: string;
				sortPosition: number;
			}>;

			for (const docConfig of additionalDocuments) {
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
		const body: IDataObject = {
			document: {
				name: mainDoc.fileName,
			},
			docContent: mainDoc.content,
			MergeDocumentsAction: {
				Documents: documentsToMerge.map(doc => ({
					FileContent: doc.FileContent,
					FileName: doc.FileName,
					SortPosition: doc.SortPosition,
				})),
			},
		};

		// Add merge options if needed
		if (maintainFormatting !== undefined) {
			(body.MergeDocumentsAction as IDataObject).MaintainFormatting = maintainFormatting;
		}

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
				const baseName = mainDoc.fileName
					? mainDoc.fileName.replace(/\.[^.]*$/, '')
					: 'merged_document';
				fileName = `${baseName}_merged.docx`;
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
						maintainFormatting,
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

