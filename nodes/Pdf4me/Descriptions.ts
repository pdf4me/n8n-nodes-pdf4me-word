/* eslint-disable n8n-nodes-base/node-filename-against-convention, n8n-nodes-base/node-param-default-missing */
import { INodeTypeDescription, NodeConnectionType } from 'n8n-workflow';
import * as addTextWatermarkToWord from './actions/addTextWatermarkToWord';
import * as addImageWatermarkToWord from './actions/addImageWatermarkToWord';
import * as extractWordMetadata from './actions/extractWordMetadata';
import * as optimizeWordDocument from './actions/optimizeWordDocument';
import * as compareWordDocuments from './actions/compareWordDocuments';
import * as splitWordDocument from './actions/splitWordDocument';
import * as mergeWordDocuments from './actions/mergeWordDocuments';
import * as secureWordDocument from './actions/secureWordDocument';
import * as deletePagesFromWord from './actions/deletePagesFromWord';
import * as updateToc from './actions/updateToc';
import * as replaceText from './actions/replaceText';
import * as updateHeadersFooters from './actions/updateHeadersFooters';
import * as replaceTextWithImage from './actions/replaceTextWithImage';
import { ActionConstants } from './GenericFunctions';

export const descriptions: INodeTypeDescription = {
	displayName: 'PDF4me Word',
	name: 'pdf4meWord',
	description: 'Process Word documents with PDF4ME\'s powerful Word processing capabilities. Add customizable text watermarks to Word documents with full control over styling, orientation, and positioning.',
	defaults: {
		name: 'PDF4me Word',
	},
	group: ['transform'],
	icon: 'file:300.svg',
	inputs: [NodeConnectionType.Main],
	outputs: [NodeConnectionType.Main],
	credentials: [
		{
			name: 'pdf4meApi',
			required: true,
		},
	], // eslint-disable-line n8n-nodes-base/node-param-default-missing
	properties: [
		{
			displayName: 'Operation',
			name: 'operation',
			type: 'options',
			noDataExpression: true,
			options: [
				{
					name: 'Add Text Watermark To Word',
					description: 'Add customizable text watermark to Word documents with font, color, rotation, and orientation options',
					value: ActionConstants.AddTextWatermarkToWord,
					action: 'Add text watermark to Word document',
				},
				{
					name: 'Add Image Watermark To Word',
					description: 'Add image watermark to Word documents with configurable scale, size, alignment, and transparency options',
					value: ActionConstants.AddImageWatermarkToWord,
					action: 'Add image watermark to Word document',
				},
				{
					name: 'Extract Word Metadata',
					description: 'Extract metadata and document properties from Word documents including statistics, properties, and other information',
					value: ActionConstants.ExtractWordMetadata,
					action: 'Extract metadata from Word document',
				},
				{
					name: 'Optimize Word Document',
					description: 'Optimize Word documents by reducing file size and improving performance with configurable optimization levels (Low, Medium, High)',
					value: ActionConstants.OptimizeWordDocument,
					action: 'Optimize Word document',
				},
				{
					name: 'Compare Word Documents',
					description: 'Compare two Word documents and create a marked-up document showing differences with configurable comparison options',
					value: ActionConstants.CompareWordDocuments,
					action: 'Compare Word documents',
				},
				{
					name: 'Split Word Document',
					description: 'Split Word documents by pages, sections, headings, or custom ranges into multiple separate documents',
					value: ActionConstants.SplitWordDocument,
					action: 'Split Word document',
				},
				{
					name: 'Merge Word Documents',
					description: 'Merge multiple Word documents into a single document with configurable merge options',
					value: ActionConstants.MergeWordDocuments,
					action: 'Merge Word documents',
				},
				{
					name: 'Delete Pages From Word',
					description: 'Delete specified pages or ranges from a Word document with options to update pagination',
					value: ActionConstants.DeletePagesFromWord,
					action: 'Delete pages from Word document',
				},
				{
					name: 'Secure Word Document',
					description: 'Apply password protection and protection types (ReadOnly, CommentsOnly, FormsOnly) to Word documents',
					value: ActionConstants.SecureWordDocument,
					action: 'Secure Word document',
				},
				{
					name: 'Update Table of Contents',
					description: 'Update table of contents in Word documents with configurable heading levels, page numbers, and tab leaders',
					value: ActionConstants.UpdateToc,
					action: 'Update table of contents',
				},
				{
					name: 'Replace Text',
					description: 'Replace text in Word documents with configurable search options, formatting, and regular expressions',
					value: ActionConstants.ReplaceText,
					action: 'Replace text in Word document',
				},
				{
					name: 'Update Headers and Footers',
					description: 'Update headers and footers in Word documents with configurable content for different page types',
					value: ActionConstants.UpdateHeadersFooters,
					action: 'Update headers and footers in Word document',
				},
				{
					name: 'Replace Text With Image',
					description: 'Replace text in Word documents with images, with configurable size, aspect ratio, and page filtering options',
					value: ActionConstants.ReplaceTextWithImage,
					action: 'Replace text with image in Word document',
				},
			],
			default: ActionConstants.AddTextWatermarkToWord,
		},
		...addTextWatermarkToWord.description,
		...addImageWatermarkToWord.description,
		...extractWordMetadata.description,
		...optimizeWordDocument.description,
		...compareWordDocuments.description,
		...splitWordDocument.description,
		...mergeWordDocuments.description,
		...secureWordDocument.description,
		...deletePagesFromWord.description,
		...updateToc.description,
		...replaceText.description,
		...updateHeadersFooters.description,
		...replaceTextWithImage.description,
	],
	subtitle: '={{$parameter["operation"]}}',
	version: 1,
};
