package graph

import (
	"encoding/json"
	"fmt"
	"net/http"
	"strings"

	api "github.com/fossteams/teams-api/pkg"
)

const defaultBaseURL = "https://graph.microsoft.com"

type Event struct {
	ID      string `json:"id"`
	Subject string `json:"subject"`
}

type eventsResponse struct {
	Value []Event `json:"value"`
}

type CalendarClient struct {
	baseURL string
	client  *http.Client
	token   *api.TeamsToken
}

func NewCalendarClient(client *http.Client, token *api.TeamsToken) *CalendarClient {
	return NewCalendarClientWithBaseURL(defaultBaseURL, client, token)
}

func NewCalendarClientWithBaseURL(baseURL string, client *http.Client, token *api.TeamsToken) *CalendarClient {
	if client == nil {
		client = http.DefaultClient
	}

	return &CalendarClient{
		baseURL: strings.TrimRight(baseURL, "/"),
		client:  client,
		token:   token,
	}
}

func (c *CalendarClient) ListMyEvents() ([]Event, error) {
	if c.token == nil || c.token.Inner == nil {
		return nil, fmt.Errorf("graph token cannot be nil")
	}

	req, err := http.NewRequest(http.MethodGet, c.baseURL+"/v1.0/me/calendar/events", nil)
	if err != nil {
		return nil, fmt.Errorf("unable to create Graph calendar request: %v", err)
	}
	req.Header.Set("Authorization", api.AuthString(c.token))

	resp, err := c.client.Do(req)
	if err != nil {
		return nil, fmt.Errorf("unable to perform Graph calendar request: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("unexpected Graph calendar status: %s", resp.Status)
	}

	var decoded eventsResponse
	if err := json.NewDecoder(resp.Body).Decode(&decoded); err != nil {
		return nil, fmt.Errorf("unable to decode Graph calendar response: %v", err)
	}

	return decoded.Value, nil
}
