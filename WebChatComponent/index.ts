import {IInputs, IOutputs} from "./generated/ManifestTypes";

interface DirectLineDetails {
    directLineURL: string;
    token: string;
}

export class WebChatComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {    
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;

    private _isInitialised: boolean = false;
    private _webChatElement: HTMLDivElement;
    private _styleOptions: any;
    private _fontSize: string = "12px";
    private _textAlign: string = "left";
    private _botTokenEndpoint: string = "";
    private _directLineDetails: DirectLineDetails = {token:'', directLineURL:''};

    private _token : string = '';
    private _directLineURL : string = '';

    constructor()
    {
    }

    private loadWebChatScript()
    {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdn.botframework.com/botframework-webchat/latest/webchat.js';
            script.onload = resolve;
            script.onerror = reject;
            document.head.appendChild(script);
        });
    }

    // private async fetchDirectLineURL(apiVersion: string, tokenEndpointURL: URL): Promise<string> {
    //     try {
    //         const response = await fetch(new URL(`/powervirtualagents/regionalchannelsettings?api-version=${apiVersion}`, tokenEndpointURL.toString()));
    //         if (!response.ok) {
    //             throw new Error('Failed to retrieve regional channel settings.');
    //         }
    //         const { channelUrlsById: { directline } } = await response.json();
    //         return directline;
    //     } catch (error) {
    //         console.error(`WebChatComponent, error getting Direct Line URL: ${JSON.stringify(error instanceof Error ? error.message : error)}`);
    //         return ''; 
    //     }
    // }

    private async fetchDirectLineToken(tokenEndpointURL: URL): Promise<string> {
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

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        // Add control initialization code
        this._container = container;
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._context.mode.trackContainerResize(true);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void
    {

        try {
            this._styleOptions = JSON.parse(this._context.parameters.styleOptions.raw || '{}');
            console.info('WebChatComponent: Style Options:', JSON.stringify(this._styleOptions));
        } catch (error : any) {
            console.error('Error parsing JSON:', error);
            this._styleOptions = {}; 
        }

        if(this._botTokenEndpoint != this._context.parameters.botTokenEndpoint.raw) 
        {
            this._botTokenEndpoint = this._context.parameters.botTokenEndpoint.raw || '';
            console.info('WebChatComponent: Bot Token Endpoint:', this._botTokenEndpoint);
        }

        if(this._fontSize != this._context.parameters.fontSize.raw)
        {
            this._fontSize = this._context.parameters.fontSize.raw || '12px';
            console.info('WebChatComponent: Font Size:', this._fontSize);
        }

        if(this._textAlign != this._context.parameters.textAlign.raw)
        {
            this._textAlign = this._context.parameters.textAlign.raw || 'left';
            console.info('WebChatComponent: Text Align:', this._textAlign);
        }

        if(this._webChatElement != null)
        {
            this._webChatElement.style.width = `${this._context.mode.allocatedWidth}px`;
            this._webChatElement.style.height = `${this._context.mode.allocatedHeight}px`;
            this._webChatElement.style.textAlign = this._textAlign;
            this._webChatElement.style.fontSize = this._fontSize;
        }

        if(this._botTokenEndpoint == '' || this._botTokenEndpoint == 'val')
        {
            console.info('WebChatComponent: Bot Token Endpoint not set');
            return;
        }

        if(!this._isInitialised)
        {
            this.loadWebChatScript().then(async () => {
                console.info('WebChatComponent: Web Chat script loaded');
                this._webChatElement = document.createElement('div');
                this._webChatElement.id = 'webchat';
                this._webChatElement.role = 'main';
                this._webChatElement.style.width = `${this._context.mode.allocatedWidth}px`;
                this._webChatElement.style.height = `${this._context.mode.allocatedHeight}px`;
                this._webChatElement.style.fontSize = this._fontSize;
                this._webChatElement.style.textAlign = this._textAlign;
                this._container.appendChild(this._webChatElement);
                console.info('WebChatComponent: Bot Token Endpoint:', this._botTokenEndpoint);
                const tokenEndpointURL = new URL(this._botTokenEndpoint);
                const locale = document.documentElement.lang || 'en';
                const apiVersion = tokenEndpointURL.searchParams.get('api-version') || "";

                //this._directLineURL = await this.fetchDirectLineURL(apiVersion, tokenEndpointURL);
                //console.info(`WebChatComponent: Direct Line URL: ${this._directLineURL}`);

                this._token = await this.fetchDirectLineToken(tokenEndpointURL);
                console.info(`WebChatComponent: Token: ${this._token}`);

                const directLineURL = this._directLineURL
                const token = this._token;

                const directLine = (window as any).WebChat.createDirectLine({ domain: new URL('v3/directline', 'https://europe.directline.botframework.com'), token });

                const subscription = directLine.connectionStatus$.subscribe({
                    next(value : any) {
                      if (value === 2) {
                        directLine
                          .postActivity({
                            localTimezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
                            locale,
                            name: 'startConversation',
                            type: 'event'
                          })
                          .subscribe();
                        // Only send the event once, unsubscribe after the event is sent.
                        subscription.unsubscribe();
                      }
                    }
                  });
                  console.info('WebChatComponent: Direct Line created and subscribed');
                  const styleOptions = this._styleOptions;
                  (window as any).WebChat.renderWebChat({ directLine, locale, styleOptions }, document.getElementById('webchat'));
                  console.info('WebChatComponent: Web Chat rendered');
                  this._isInitialised = true;
            });
        } 
    }

    public getOutputs(): IOutputs
    {
        return {};
    }

    public destroy(): void
    {
    }
}
