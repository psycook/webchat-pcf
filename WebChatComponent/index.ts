import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class WebChatComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;

    private _isScriptLoaded: boolean = false;
    private _isScriptLoading: boolean = false;
    private _webChatElement: HTMLDivElement;
    private _styleOptions: any;
    private _fontSize: string = "12px";
    private _textAlign: string = "left";
    private _botTokenEndpoint: string = "";
    private _locale: string = 'en';
    private _token: string = '';
    private _directLine: any;
    private _subscriptionNeedsUpdate: boolean = false;
    private _needsRerender: boolean = false;
    private _lastKnownWidth: number;
    private _lastKnownHeight: number;

    private _DEBUG: boolean = false;

    constructor() {
    }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._container = container;
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._context.mode.trackContainerResize(true);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.checkForChanges(context);
        if (!this._isScriptLoaded && !this._isScriptLoading) {
            this.loadWebChatScript().then(() => {
                this.initializeWebChat();
            });
        } else if (this._needsRerender) {
            this.initializeWebChat();
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
    }

    private async initializeWebChat(): Promise<void> {
        if(this._DEBUG) console.log('WebChatComponent: Initializing WebChat');
        if (!this._webChatElement) {
            this.createComponent();
        } else {
            this.updateComponent();
        }
        if (this._token === '' || this._subscriptionNeedsUpdate) {
            try {
                const tokenEndpointURL = new URL(this._botTokenEndpoint);
                this._token = await this.fetchDirectLineToken(tokenEndpointURL);
                if(this._DEBUG) console.log(`WebChatComponent: Direct Line token updated to ${this._token}`);
            } catch (error) {
                if(this._DEBUG) console.error('Failed to fetch Direct Line token:', error);
                return; // Optionally handle error-specific logic, like notifying the user
            }
            this.subscribe();
            this._subscriptionNeedsUpdate = false;
        }
    
        if (this._needsRerender) {
            this.renderWebChat();
            this._needsRerender = false;
        }
    }
    
    
    private checkForChanges(context: ComponentFramework.Context<IInputs>): void {
        const newTokenEndpoint = context.parameters.botTokenEndpoint.raw || '';
        const newStyleOptions = JSON.parse(context.parameters.styleOptions.raw || '{}');
        const newFontSize = context.parameters.fontSize.raw || '12px';
        const newTextAlign = context.parameters.textAlign.raw || 'left';

        if (this._lastKnownWidth !== context.mode.allocatedWidth || this._lastKnownHeight !== context.mode.allocatedHeight) {
            this._lastKnownWidth = context.mode.allocatedWidth;
            this._lastKnownHeight = context.mode.allocatedHeight;
    
            // Update web chat dimensions
            if (this._webChatElement) {
                this.updateComponentSize();
            }
        }
    
        if (this._botTokenEndpoint !== newTokenEndpoint) {
            this._botTokenEndpoint = newTokenEndpoint;
            this._subscriptionNeedsUpdate = true;
        }
    
        if (JSON.stringify(this._styleOptions) !== JSON.stringify(newStyleOptions)) {
            this._styleOptions = newStyleOptions;
            this._needsRerender = true;
        }
    
        if (this._fontSize !== newFontSize || this._textAlign !== newTextAlign) {
            this._fontSize = newFontSize;
            this._textAlign = newTextAlign;
            this._needsRerender = true;
        }
    }

    private loadWebChatScript() 
    {
        if(this._DEBUG) console.log('WebChatComponent: Loading WebChat script');
        this._isScriptLoading = true;
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdn.botframework.com/botframework-webchat/latest/webchat.js';
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    private createComponent() 
    {
        if(this._DEBUG) console.log('WebChatComponent: Creating WebChat component');
        this._webChatElement = document.createElement('div');
        this._webChatElement.id = 'webchat';
        this._webChatElement.role = 'main';
        this._webChatElement.style.width = `${this._context.mode.allocatedWidth}px`;
        this._webChatElement.style.height = `${this._context.mode.allocatedHeight}px`;
        this._webChatElement.style.fontSize = `${this._fontSize}`;
        this._webChatElement.style.textAlign = this._textAlign;
        this._container.appendChild(this._webChatElement);
    }

    private updateComponent() 
    {
        if(this._DEBUG) console.log('WebChatComponent: Updating WebChat component');
        this._webChatElement.style.fontSize = this._fontSize;
        this._webChatElement.style.textAlign = this._textAlign;
        this.renderWebChat();
    }

    private updateComponentSize(): void {
        if(this._DEBUG) console.log('WebChatComponent: Updating WebChat component size');
        if (this._webChatElement) {
            this._webChatElement.style.width = `${this._lastKnownWidth}px`;
            this._webChatElement.style.height = `${this._lastKnownHeight}px`;
            this.renderWebChat(); // Rerender or refresh the chat interface if necessary
        }
    }

    private subscribe() 
    {
        if(this._DEBUG) console.log('WebChatComponent: Subscribing to Direct Line');
        const token = this._token;
        const locale = this._locale;
        this._directLine = (window as any).WebChat.createDirectLine(
            { domain: new URL('v3/directline', 'https://europe.directline.botframework.com'), token }
        );
        const directLine = this._directLine;
        const subscription = directLine.connectionStatus$.subscribe(
            {
                next(value: any) {
                    if (value === 2) {
                        directLine.postActivity(
                            {
                                localTimezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
                                locale,
                                name: 'startConversation',
                                type: 'event'
                            }
                        ).subscribe();
                        subscription.unsubscribe();
                    }
                }
            }
        );
        this.renderWebChat();
    }

    private renderWebChat() 
    {
        if(this._DEBUG) console.log('WebChatComponent: Rendering WebChat');
        const directLine = this._directLine;
        const locale = this._locale;
        const styleOptions = this._styleOptions;

        (window as any).WebChat.renderWebChat(
            { 
                directLine, 
                locale, 
                styleOptions 
            },
            document.getElementById('webchat')
        );
    }

    private async fetchDirectLineToken(tokenEndpointURL: URL): Promise<string> 
    {
        if(this._DEBUG) console.log(`WebChatComponent:fetchDirectLineToken() Fetching Direct Line token from ${tokenEndpointURL.toString()}`);
        try {
            const response = await fetch(tokenEndpointURL.toString());
            if (!response.ok) {
                throw new Error('Failed to retrieve Direct Line token.');
            }
            const { token } = await response.json();
            return token;
        } catch (error) {
            console.error(`WebChatComponent, error getting token direct line: ${JSON.stringify(error instanceof Error ? error.message : error)}`);
            return '';
        }
    }
}