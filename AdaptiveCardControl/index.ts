import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as AdaptiveCards from "adaptivecards";
import * as ACData from "adaptivecards-templating";

export class AdaptiveCardControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _notifyOutputChanged: () => void;
	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _card: HTMLElement;
	private _responseData: string;

	constructor() { }

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this._notifyOutputChanged = notifyOutputChanged;
		this._context = context;
		this._container = container;
	}

	public async updateView(context: ComponentFramework.Context<IInputs>): Promise<void> {

		this._context = context;

		if (this._card) {
			this._container.removeChild(this._card);
		}

		if (this._context.updatedProperties.includes("responseJson")) {
			this.renderResponse('Thank you for submitting!')
		}
		else if (this._context.parameters.responseJson.raw) {
			this.renderResponse('Thank you for submitting!');
		}
		else {
			const adaptiveCardTemplateJson = await this.getAdaptiveCardJson(
				this._context.parameters.templateRetrieveEntityType.raw,
				this._context.parameters.templateRetrieveOptions.raw,
				this._context.parameters.templateRetrieveAttributeName.raw,
				this._context.parameters.templateRetrieveFetchXml.raw);

			const adaptiveCardDataJson = await this.getAdaptiveCardJson(
				this._context.parameters.dataRetrieveEntityType.raw,
				this._context.parameters.dataRetrieveOptions.raw,
				this._context.parameters.dataRetrieveAttributeName.raw,
				this._context.parameters.dataRetrieveFetchXml.raw);

			this.renderRequestCard(adaptiveCardTemplateJson, adaptiveCardDataJson);
		}
	}

	private async getAdaptiveCardJson(entityType: string, options: string, attributeName: string, fetchXml?: string): Promise<any> {

		// @ts-ignore
		const entityId = this._context.page.entityId;
		const retrieveOptions = (fetchXml ? "?fetchXml=" + fetchXml : options || "").replace("{entityId}", entityId);
		const response = await this._context.webAPI.retrieveMultipleRecords(entityType, retrieveOptions);

		if (response && response.entities.length == 1) {
			const adaptiveCardJson = response.entities[0][attributeName];
			const json = JSON.parse(adaptiveCardJson);
			return json;
		}

		return null;
	}

	private renderRequestCard(templateJson: any, dataJson: any) {
		if (!templateJson || !dataJson) return;

		const template = new ACData.Template(templateJson);
		const cardPayload = template.expand({
			$root: dataJson
		});

		const adaptiveCard = new AdaptiveCards.AdaptiveCard();
		adaptiveCard.parse(cardPayload);

		adaptiveCard.onExecuteAction = this.submitRequest.bind(this);

		this._card = adaptiveCard.render();
		this._container.appendChild(this._card);
	}

	private renderResponse(message: string) {
		this._card = document.createElement("div");
		this._card.innerHTML = `<div>${message}</div>`;
		this._container.appendChild(this._card);
	}

	private submitRequest(action: AdaptiveCards.Action) {
		const submitAction = action as AdaptiveCards.SubmitAction;
		if (submitAction) {
			this._responseData = JSON.stringify(submitAction.data);
			this._notifyOutputChanged();
		}
	}

	public getOutputs(): IOutputs {
		return {
			responseJson: this._responseData
		};
	}

	public destroy(): void {
	}
}