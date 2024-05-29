package msgraph

import (
	"context"
	"errors"
	"time"

	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	abstractions "github.com/microsoft/kiota-abstractions-go"
	azureauth "github.com/microsoft/kiota-authentication-azure-go"
	http "github.com/microsoft/kiota-http-go"
	msgraphsdk "github.com/microsoftgraph/msgraph-sdk-go"
	"github.com/microsoftgraph/msgraph-sdk-go/models"
	"github.com/microsoftgraph/msgraph-sdk-go/models/odataerrors"
	"github.com/microsoftgraph/msgraph-sdk-go/users"

	nethttp "net/http"
)

// Credentials client credentials flow
type Credentials struct {
	ClientID     string
	ClientSecret string
	TenantID     string
}

// Service is a struct that holds the authentication credentials and the GraphServiceClient.
type Service struct {
	auth  Credentials
	graph msgraphsdk.GraphServiceClient
}

// NewService creates a new instance of the Service struct.
// It uses the Azure Identity library to create a new client secret credential.
// It then creates a new Azure Identity Authentication Provider with the required scopes.
// Finally, it creates a new GraphServiceClient with the authentication provider and returns it.
// It takes a Credentials struct as input and returns a pointer to a Service struct and an error.
func NewService(c Credentials) (*Service, error) {
	credentials, err := azidentity.NewClientSecretCredential(
		c.TenantID,
		c.ClientID,
		c.ClientSecret,
		nil,
	)
	if err != nil {
		return nil, parseError(err)
	}

	auth, err := azureauth.NewAzureIdentityAuthenticationProviderWithScopes(credentials, []string{"https://graph.microsoft.com/.default"})
	if err != nil {
		return nil, parseError(err)
	}

	httpClient := &nethttp.Client{
		Timeout: time.Minute * 1,
	}
	ra, err := http.NewNetHttpRequestAdapterWithParseNodeFactoryAndSerializationWriterFactoryAndHttpClient(auth, nil, nil, httpClient)
	if err != nil {
		return nil, parseError(err)
	}

	return &Service{
		auth:  c,
		graph: *msgraphsdk.NewGraphServiceClient(ra),
	}, nil
}

// GetMailFolderMessagesDeltaLink is a method on the Service struct.
// It uses the GraphServiceClient to create a request to get the delta link for the messages in the specified mail folder.
// It then sends the request and returns the delta link.
// It takes a context, a user ID, and a mail folder ID as input.
// It returns a pointer to a string and an error.
func (c *Service) GetMailFolderMessagesDeltaLink(ctx context.Context, userId string, mailFolderId string) (*string, error) {
	requestBuilder := c.graph.UsersById(userId).MailFoldersById(mailFolderId).Messages().MicrosoftGraphDelta()
	ri, err := requestBuilder.ToGetRequestInformation(ctx, nil)
	if err != nil {
		return nil, parseError(err)
	}

	ri.UrlTemplate = "{+baseurl}/users/{user%2Did}/mailFolders/{mailFolder%2Did}/messages/microsoft.graph.delta()?changeType=created{?%24top,%24skip,%24search,%24filter,%24count,%24select,%24orderby}"
	errorMapping := abstractions.ErrorMappings{
		"4XX": odataerrors.CreateODataErrorFromDiscriminatorValue,
		"5XX": odataerrors.CreateODataErrorFromDiscriminatorValue,
	}

	response, err := c.graph.GetAdapter().Send(ctx, ri, users.CreateItemMailFoldersItemMessagesMicrosoftGraphDeltaDeltaResponseFromDiscriminatorValue, errorMapping)
	if err != nil {
		return nil, parseError(err)
	}
	result := response.(users.ItemMailFoldersItemMessagesMicrosoftGraphDeltaDeltaResponseable)

	//skip existing messages
	for result.GetOdataNextLink() != nil {
		result, err = users.NewItemMailFoldersItemMessagesMicrosoftGraphDeltaRequestBuilder(*result.GetOdataNextLink(), c.graph.GetAdapter()).Get(ctx, nil)
		if err != nil {
			return nil, parseError(err)
		}
	}

	return result.GetOdataDeltaLink(), nil
}

// GetMessagesDelta is a method on the Service struct.
// It uses the GraphServiceClient to create a request to get the messages from the delta link.
// It then sends the request and returns the messages and the delta link.
// It takes a context and a delta link as input.
// It returns a slice of Messageable, a string, and an error.
func (c *Service) GetMessagesDelta(ctx context.Context, deltaLink string) ([]models.Messageable, string, error) {
	var result []models.Messageable
	response, err := users.NewItemMailFoldersItemMessagesMicrosoftGraphDeltaRequestBuilder(deltaLink, c.graph.GetAdapter()).Get(ctx, nil)
	if err != nil {
		return nil, "", parseError(err)
	}
	result = append(result, response.GetValue()...)

	for response.GetOdataDeltaLink() == nil {
		response, err = users.NewItemMailFoldersItemMessagesMicrosoftGraphDeltaRequestBuilder(*response.GetOdataNextLink(), c.graph.GetAdapter()).Get(ctx, nil)
		if err != nil {
			return nil, "", parseError(err)
		}
		result = append(result, response.GetValue()...)
	}

	dl := *response.GetOdataDeltaLink()

	return result, dl, nil
}

// GetAttachments is a method on the Service struct.
// It uses the GraphServiceClient to create a request to get the attachments of the specified message.
// It then sends the request and returns the attachments.
// It takes a context, a user ID, a message ID, and a boolean indicating whether to include the content of the attachments as input.
// It returns a slice of FileAttachment and an error.
func (c *Service) GetAttachments(ctx context.Context, userId string, messageId string, withContent bool) ([]FileAttachment, error) {
	result, err := c.graph.UsersById(userId).MessagesById(messageId).Attachments().Get(ctx, nil)
	if err != nil {
		return nil, parseError(err)
	}

	var attachments []FileAttachment
	for _, att := range result.GetValue() {
		if fileAtt, ok := att.(models.FileAttachmentable); ok {
			attachment := FileAttachment{
				Name:        *fileAtt.GetName(),
				ContentType: *fileAtt.GetContentType(),
			}
			if withContent {
				attachment.Content = fileAtt.GetContentBytes()
			}

			attachments = append(attachments, attachment)
		}
	}

	return attachments, nil
}

// GetMessage is a method on the Service struct.
// It uses the GraphServiceClient to create a request to get the specified message.
// It then sends the request and returns the message.
// It takes a context, a user ID, and a message ID as input.
// It returns a Messageable and an error.
func (c *Service) GetMessage(ctx context.Context, userId string, messageId string) (models.Messageable, error) {
	result, err := c.graph.UsersById(userId).MessagesById(messageId).Get(ctx, nil)
	if err != nil {
		return nil, parseError(err)
	}

	return result, nil
}

// SendMessage is a method on the Service struct.
// It uses the GraphServiceClient to create a request to send a new message.
// It then sends the request.
// It takes a context, a recipient email, a sender email, a subject, and a content as input.
// It returns an error.
func (c *Service) SendMessage(ctx context.Context, to string, from string, subject string, content string) error {
	requestBody := users.NewItemMicrosoftGraphSendMailSendMailPostRequestBody()
	message := models.NewMessage()
	message.SetSubject(&subject)

	ct := models.HTML_BODYTYPE
	body := models.NewItemBody()
	body.SetContentType(&ct)
	body.SetContent(&content)
	message.SetBody(body)

	recipient := models.NewRecipient()
	emailAddress := models.NewEmailAddress()
	emailAddress.SetAddress(&to)
	recipient.SetEmailAddress(emailAddress)

	toRecipients := []models.Recipientable{
		recipient,
	}
	message.SetToRecipients(toRecipients)
	requestBody.SetMessage(message)

	err := c.graph.UsersById(from).MicrosoftGraphSendMail().Post(ctx, requestBody, nil)
	return parseError(err)
}

// parseError is a helper function.
// It checks if the input error is an ODataError.
// If it is, it returns a new error with the message from the ODataError.
// If it is not, it returns a new error with the message from the input error.
// It takes an error as input and returns an error.
func parseError(err error) error {
	if err == nil {
		return nil
	}

	var ode *odataerrors.ODataError
	if errors.As(err, &ode) {
		if err := ode.GetError(); err != nil {
			if err.GetMessage() != nil {
				return errors.New(*err.GetMessage())
			}
		}
	}

	return errors.New(err.Error())
}

// FileAttachment is a struct that holds the name, content type, and content of a file attachment.
type FileAttachment struct {
	Name        string
	ContentType string
	Content     []byte
}
