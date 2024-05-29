package main

import (
	"context"
	"fmt"
	"os"

	"github.com/philous/office-365-listener/msgraph"
)

func main() {
	// Get configuration values from environment variables
	clientID := os.Getenv("CLIENT_ID")
	clientSecret := os.Getenv("CLIENT_SECRET")
	tenantID := os.Getenv("TENANT_ID")
	userId := os.Getenv("USER_ID")
	messageId := os.Getenv("MESSAGE_ID")

	// Create a new instance of the Service struct
	service, err := msgraph.NewService(msgraph.Credentials{
		ClientID:     clientID,
		ClientSecret: clientSecret,
		TenantID:     tenantID,
	})
	if err != nil {
		fmt.Printf("Error creating service: %v\n", err)
		return
	}

	// Use the service to get the specified message
	message, err := service.GetMessage(context.Background(), userId, messageId)
	if err != nil {
		fmt.Printf("Error getting message: %v\n", err)
		return
	}

	// Print out the details of the message
	fmt.Printf("Subject: %s\n", *message.GetSubject())
	fmt.Printf("Body: %s\n", *message.GetBody().GetContent())
	fmt.Printf("Received: %s\n", message.GetReceivedDateTime().Format("2006-01-02 15:04:05"))
}
