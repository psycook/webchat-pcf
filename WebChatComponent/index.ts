import { IInputs, IOutputs } from "./generated/ManifestTypes";

const DEFAULT_LOCALE = 'en';
const SCRIPT_SRC = 'https://cdn.botframework.com/botframework-webchat/latest/webchat.js';
const WEBCHAT_ELEMENT_ID = 'webchat';
const WEBCHAT_ELEMENT_ROLE = 'main';
const DIRECTLINE_VERSION = 'v3/directline';
const DIRECTLINE_URL = 'https://europe.directline.botframework.com';
const CONVERSATION_EVENT_NAME = 'startConversation';
const CONVERSATION_EVENT_TYPE = 'event';

export class WebChatComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _isScriptLoaded: boolean = false;
    private _isScriptLoading: boolean = false;
    private _webChatElement: HTMLDivElement;
    private _styleOptions: any;
    private _fontSize: string;
    private _textAlign: string;
    private _botTokenEndpoint: string;
    private _token: string;
    private _directLine: any;
    private _needsRerender: boolean = false;
    private _needsUpdate : boolean = false;
    private _lastKnownWidth: number;
    private _lastKnownHeight: number;
    private _debug: boolean = false;

    constructor() {
    }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._container = container;
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._context.mode.trackContainerResize(true);
        this.createComponent();
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.checkForChanges(context);
    
        if(this._botTokenEndpoint === "" || this._botTokenEndpoint === "val") return;
    
        if (!this._isScriptLoaded && !this._isScriptLoading) {
            this.loadWebChatScript().then(async () => {
                if(this._debug) console.log('WebChatComponent:updateView() WebChat script loaded');
                try {
                    // Get the token using await
                    const tokenEndpointURL = new URL(this._botTokenEndpoint);
                    if(this._debug) console.log(`WebChatComponent:updateView() Getting token.`);
                    this._token = await this.fetchDirectLineToken(tokenEndpointURL);
                    if(this._debug) console.log(`WebChatComponent:updateView() Direct Line token updated to ${this._token}`);
                    this.subscribe();
                } catch (error) {
                    if(this._debug) console.error('WebChatComponent:updateView() Error fetching Direct Line token', error);
                }
                // Create the subscription
                this.checkForUpdates();
            }).catch((error) => {
                if(this._debug) console.error('WebChatComponent:updateView() Error loading WebChat script', error);
            });
        } else if (this._isScriptLoaded) {
            this.checkForUpdates();
        }
    }
    

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
    }

    private checkForChanges(context: ComponentFramework.Context<IInputs>): void {
        const newTokenEndpoint = context.parameters.botTokenEndpoint.raw || '';
        const newStyleOptions = JSON.parse(context.parameters.styleOptions.raw || '{}');
        const newFontSize = context.parameters.fontSize.raw || '12px';
        const newTextAlign = context.parameters.textAlign.raw || 'left';
        const debug = context.parameters.enableDebug.raw || false;

        if (this._botTokenEndpoint !== newTokenEndpoint) {
            this._botTokenEndpoint = newTokenEndpoint;
        }

        if(this._debug !== debug) {
            this._debug = debug;
        }

        if (this._lastKnownWidth !== context.mode.allocatedWidth || this._lastKnownHeight !== context.mode.allocatedHeight) {
            this._lastKnownWidth = context.mode.allocatedWidth;
            this._lastKnownHeight = context.mode.allocatedHeight;
            this._needsUpdate = true;
        }

        if (this._fontSize !== newFontSize || this._textAlign !== newTextAlign) {
            this._fontSize = newFontSize;
            this._textAlign = newTextAlign;
            this._needsUpdate = true;
        }
    
        if (JSON.stringify(this._styleOptions) !== JSON.stringify(newStyleOptions)) {
            this._styleOptions = newStyleOptions;
            this._needsRerender = true;
        }
    }

    private checkForUpdates(): void {
        if(this._debug) console.log('WebChatComponent:checkForUpdates() Checking for updates');

        if (this._needsUpdate) {
            this.updateComponent();
            this._needsUpdate = false;
        }
    
        if (this._needsRerender) {
            this.renderWebChat();
            this._needsRerender = false;
        }
    }

    private loadWebChatScript(): Promise<void> {
        this._isScriptLoading = true;
        return new Promise<void>((resolve, reject) => {
            const script = document.createElement('script');
            script.src = SCRIPT_SRC;
            script.onload = () => {
                this._isScriptLoaded = true;
                this._isScriptLoading = false;
                resolve();
            };
            script.onerror = (error) => {
                this._isScriptLoaded = false;
                this._isScriptLoading = false;
                console.error('WebChatComponent:loadWebChatScript() Failed to load WebChat script', error);
                reject(new Error('Failed to load script'));
            };
            document.head.appendChild(script);
        });
    }
    
    private createComponent() 
    {
        if(this._debug) console.log('WebChatComponent:createComponent() Creating WebChat component');
        this._webChatElement = document.createElement('div');
        this._webChatElement.id = WEBCHAT_ELEMENT_ID;
        this._webChatElement.role = WEBCHAT_ELEMENT_ROLE;
        this._container.appendChild(this._webChatElement);
        this.updateComponent();
    }

    private updateComponent() 
    {
        if(this._debug) console.log('WebChatComponent:updateComponent() Updating WebChat component');
        if (this._webChatElement) {
            this._webChatElement.style.width = `${this._lastKnownWidth}px`;
            this._webChatElement.style.height = `${this._lastKnownHeight}px`;
            this._webChatElement.style.fontSize = this._fontSize;
            this._webChatElement.style.textAlign = this._textAlign;
            // this.renderWebChat(); 
        }
        else 
        {
            console.error('WebChatComponent:updateComponent() WebChat element not found, this should not happen.');
        }
    }

    private subscribe() 
    {
        if(this._debug) console.log('WebChatComponent:subscribe() Subscribing to Direct Line');
        const token = this._token;
        this._directLine = (window as any).WebChat.createDirectLine( { domain: new URL(DIRECTLINE_VERSION, DIRECTLINE_URL), token } );
        const directLine = this._directLine;
        const subscription = directLine.connectionStatus$.subscribe(
        {
            next(value: any) {
                if (value === 2) {
                    directLine.postActivity(
                        {
                            localTimezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
                            DEFAULT_LOCALE,
                            name: CONVERSATION_EVENT_NAME,
                            type: CONVERSATION_EVENT_TYPE
                        }
                    ).subscribe();
                    subscription.unsubscribe();
                }
            }
        });
        this.renderWebChat();
    }

    private renderWebChat() 
    {
        if(this._debug) console.log('WebChatComponent:renderWebChat() Rendering WebChat');
        const directLine = this._directLine;
        const styleOptions = this._styleOptions;

        (window as any).WebChat.renderWebChat(
            { 
                directLine, 
                DEFAULT_LOCALE, 
                styleOptions 
            },
            document.getElementById(WEBCHAT_ELEMENT_ID)
        );
    }

    private async fetchDirectLineToken(tokenEndpointURL: URL): Promise<string> 
    {
        if(this._debug) console.log(`WebChatComponent:fetchDirectLineToken() Fetching Direct Line token from ${tokenEndpointURL.toString()}`);
        try {
            const response = await fetch(tokenEndpointURL.toString());
            if (!response.ok) {
                throw new Error('Failed to retrieve Direct Line token.');
            }
            const { token } = await response.json();
            return token;
        } catch (error) {
            console.error(`WebChatComponent:fetchDirectLineToken() Error getting token direct line: ${JSON.stringify(error instanceof Error ? error.message : error)}`);
            return '';
        }
    }
}