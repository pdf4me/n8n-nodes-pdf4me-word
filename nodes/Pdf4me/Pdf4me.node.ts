import {
	IExecuteFunctions,
	INodeType,
	INodeTypeDescription,
	INodeTypeBaseDescription,
	INodeExecutionData,
} from 'n8n-workflow';

import { descriptions } from './Descriptions';
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
import { ActionConstants } from './GenericFunctions';

export class Pdf4me implements INodeType {
	description: INodeTypeDescription;

	constructor(baseDescription: INodeTypeBaseDescription) {
		this.description = {
			...baseDescription,
			...descriptions,
		};
	}

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const operationResult: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			const action = this.getNodeParameter('operation', i);

			try {
				if (action === ActionConstants.AddTextWatermarkToWord) {
					operationResult.push(...(await addTextWatermarkToWord.execute.call(this, i)));
				} else if (action === ActionConstants.AddImageWatermarkToWord) {
					operationResult.push(...(await addImageWatermarkToWord.execute.call(this, i)));
				} else if (action === ActionConstants.ExtractWordMetadata) {
					operationResult.push(...(await extractWordMetadata.execute.call(this, i)));
				} else if (action === ActionConstants.OptimizeWordDocument) {
					operationResult.push(...(await optimizeWordDocument.execute.call(this, i)));
				} else if (action === ActionConstants.CompareWordDocuments) {
					operationResult.push(...(await compareWordDocuments.execute.call(this, i)));
				} else if (action === ActionConstants.SplitWordDocument) {
					operationResult.push(...(await splitWordDocument.execute.call(this, i)));
				} else if (action === ActionConstants.MergeWordDocuments) {
					operationResult.push(...(await mergeWordDocuments.execute.call(this, i)));
				} else if (action === ActionConstants.SecureWordDocument) {
					operationResult.push(...(await secureWordDocument.execute.call(this, i)));
				} else if (action === ActionConstants.DeletePagesFromWord) {
					operationResult.push(...(await deletePagesFromWord.execute.call(this, i)));
				} else if (action === ActionConstants.UpdateToc) {
					operationResult.push(...(await updateToc.execute.call(this, i)));
				} else if (action === ActionConstants.ReplaceText) {
					operationResult.push(...(await replaceText.execute.call(this, i)));
				} else if (action === ActionConstants.UpdateHeadersFooters) {
					operationResult.push(...(await updateHeadersFooters.execute.call(this, i)));
				}
			} catch (err) {
				if (this.continueOnFail()) {
					operationResult.push({ json: this.getInputData(i)[0].json, error: err });
				} else {
					throw err;
				}
			}
		}

		return [operationResult];
	}
}
