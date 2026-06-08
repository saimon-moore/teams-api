package graph

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"

	api "github.com/saimon-moore/teams-api/pkg"
)

const defaultBaseURL = "https://graph.microsoft.com"

type EmailAddress struct {
	Name    string `json:"name,omitempty"`
	Address string `json:"address,omitempty"`
}

type Recipient struct {
	EmailAddress EmailAddress `json:"emailAddress"`
}

type Calendar struct {
	ID                string    `json:"id"`
	Name              string    `json:"name"`
	CanEdit           bool      `json:"canEdit"`
	IsDefaultCalendar bool      `json:"isDefaultCalendar"`
	Owner             EmailAddress `json:"owner"`
}

type DateTimeTimeZone struct {
	DateTime string `json:"dateTime"`
	TimeZone string `json:"timeZone,omitempty"`
}

type ItemBody struct {
	ContentType string `json:"contentType,omitempty"`
	Content     string `json:"content,omitempty"`
}

type Location struct {
	DisplayName string `json:"displayName,omitempty"`
}

type Event struct {
	ID          string           `json:"id"`
	Subject     string           `json:"subject"`
	Start       DateTimeTimeZone `json:"start"`
	End         DateTimeTimeZone `json:"end"`
	IsAllDay    bool             `json:"isAllDay"`
	Location    Location         `json:"location"`
	Body        ItemBody         `json:"body"`
	BodyPreview string           `json:"bodyPreview,omitempty"`
	WebLink     string           `json:"webLink,omitempty"`
}

type ListEventsOptions struct {
	CalendarID string
	Start      time.Time
	End        time.Time
	TimeZone   string
	Limit      int
}

type CreateEventInput struct {
	CalendarID string
	Subject    string
	Start      time.Time
	End        time.Time
	TimeZone   string
	Location   string
	Body       string
	AllDay     bool
}

type UpdateEventInput struct {
	Subject  *string
	Start    *time.Time
	End      *time.Time
	TimeZone *string
	Location *string
	Body     *string
	AllDay   *bool
}

type calendarsResponse struct {
	Value []Calendar `json:"value"`
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

func (c *CalendarClient) ListCalendars() ([]Calendar, error) {
	if err := c.validateToken(); err != nil {
		return nil, err
	}

	req, err := c.newRequest(http.MethodGet, "/v1.0/me/calendars", nil)
	if err != nil {
		return nil, err
	}

	var decoded calendarsResponse
	if err := c.doJSON(req, http.StatusOK, &decoded); err != nil {
		return nil, err
	}

	return decoded.Value, nil
}

func (c *CalendarClient) ListEvents(opts ListEventsOptions) ([]Event, error) {
	if err := c.validateToken(); err != nil {
		return nil, err
	}
	if opts.Start.IsZero() || opts.End.IsZero() {
		return nil, fmt.Errorf("list events requires start and end values")
	}

	path := "/v1.0/me/calendarView"
	if strings.TrimSpace(opts.CalendarID) != "" {
		path = "/v1.0/me/calendars/" + url.PathEscape(strings.TrimSpace(opts.CalendarID)) + "/calendarView"
	}

	req, err := c.newRequest(http.MethodGet, path, nil)
	if err != nil {
		return nil, err
	}

	query := req.URL.Query()
	query.Set("startDateTime", opts.Start.Format(time.RFC3339))
	query.Set("endDateTime", opts.End.Format(time.RFC3339))
	if opts.Limit > 0 {
		query.Set("$top", fmt.Sprintf("%d", opts.Limit))
	}
	req.URL.RawQuery = query.Encode()
	if tz := strings.TrimSpace(opts.TimeZone); tz != "" {
		req.Header.Set("Prefer", fmt.Sprintf(`outlook.timezone="%s"`, tz))
	}

	var decoded eventsResponse
	if err := c.doJSON(req, http.StatusOK, &decoded); err != nil {
		return nil, err
	}

	return decoded.Value, nil
}

func (c *CalendarClient) CreateEvent(input CreateEventInput) (*Event, error) {
	if err := c.validateToken(); err != nil {
		return nil, err
	}

	body, err := json.Marshal(buildCreatePayload(input))
	if err != nil {
		return nil, fmt.Errorf("unable to encode Graph create event payload: %v", err)
	}

	path := "/v1.0/me/events"
	if strings.TrimSpace(input.CalendarID) != "" {
		path = "/v1.0/me/calendars/" + url.PathEscape(strings.TrimSpace(input.CalendarID)) + "/events"
	}

	req, err := c.newRequest(http.MethodPost, path, bytes.NewReader(body))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Content-Type", "application/json")

	var event Event
	if err := c.doJSON(req, http.StatusCreated, &event); err != nil {
		return nil, err
	}

	return &event, nil
}

func (c *CalendarClient) UpdateEvent(calendarID, eventID string, input UpdateEventInput) (*Event, error) {
	if err := c.validateToken(); err != nil {
		return nil, err
	}

	body, err := json.Marshal(buildUpdatePayload(input))
	if err != nil {
		return nil, fmt.Errorf("unable to encode Graph update event payload: %v", err)
	}

	path := "/v1.0/me/events/" + url.PathEscape(strings.TrimSpace(eventID))
	if strings.TrimSpace(calendarID) != "" {
		path = "/v1.0/me/calendars/" + url.PathEscape(strings.TrimSpace(calendarID)) + "/events/" + url.PathEscape(strings.TrimSpace(eventID))
	}

	req, err := c.newRequest(http.MethodPatch, path, bytes.NewReader(body))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Content-Type", "application/json")

	var event Event
	if err := c.doJSON(req, http.StatusOK, &event); err != nil {
		return nil, err
	}

	return &event, nil
}

func (c *CalendarClient) DeleteEvent(calendarID, eventID string) error {
	if err := c.validateToken(); err != nil {
		return err
	}

	path := "/v1.0/me/events/" + url.PathEscape(strings.TrimSpace(eventID))
	if strings.TrimSpace(calendarID) != "" {
		path = "/v1.0/me/calendars/" + url.PathEscape(strings.TrimSpace(calendarID)) + "/events/" + url.PathEscape(strings.TrimSpace(eventID))
	}

	req, err := c.newRequest(http.MethodDelete, path, nil)
	if err != nil {
		return err
	}

	return c.doNoContent(req, http.StatusNoContent)
}

func (c *CalendarClient) ListMyEvents() ([]Event, error) {
	now := time.Now().UTC()
	return c.ListEvents(ListEventsOptions{
		Start: now,
		End:   now.Add(7 * 24 * time.Hour),
	})
}

func (c *CalendarClient) validateToken() error {
	if c.token == nil || c.token.Inner == nil {
		return fmt.Errorf("graph token cannot be nil")
	}
	return nil
}

func (c *CalendarClient) newRequest(method, path string, body io.Reader) (*http.Request, error) {
	req, err := http.NewRequest(method, c.baseURL+path, body)
	if err != nil {
		return nil, fmt.Errorf("unable to create Graph calendar request: %v", err)
	}
	req.Header.Set("Authorization", api.AuthString(c.token))
	return req, nil
}

func (c *CalendarClient) doJSON(req *http.Request, wantStatus int, target any) error {
	resp, err := c.client.Do(req)
	if err != nil {
		return fmt.Errorf("unable to perform Graph calendar request: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != wantStatus {
		return graphStatusError(resp)
	}
	if err := json.NewDecoder(resp.Body).Decode(target); err != nil {
		return fmt.Errorf("unable to decode Graph calendar response: %v", err)
	}
	return nil
}

func (c *CalendarClient) doNoContent(req *http.Request, wantStatus int) error {
	resp, err := c.client.Do(req)
	if err != nil {
		return fmt.Errorf("unable to perform Graph calendar request: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != wantStatus {
		return graphStatusError(resp)
	}
	return nil
}

func graphStatusError(resp *http.Response) error {
	bodyBytes, _ := io.ReadAll(io.LimitReader(resp.Body, 4096))
	body := strings.TrimSpace(string(bodyBytes))
	if body == "" {
		return fmt.Errorf("unexpected Graph calendar status: %s", resp.Status)
	}
	return fmt.Errorf("unexpected Graph calendar status: %s: %s", resp.Status, body)
}

func buildCreatePayload(input CreateEventInput) map[string]any {
	payload := map[string]any{
		"subject":  input.Subject,
		"isAllDay": input.AllDay,
		"start": map[string]any{
			"dateTime": input.Start.Format(time.RFC3339),
			"timeZone": normalizeTimeZone(input.TimeZone),
		},
		"end": map[string]any{
			"dateTime": input.End.Format(time.RFC3339),
			"timeZone": normalizeTimeZone(input.TimeZone),
		},
	}

	if location := strings.TrimSpace(input.Location); location != "" {
		payload["location"] = map[string]any{"displayName": location}
	}
	if body := strings.TrimSpace(input.Body); body != "" {
		payload["body"] = map[string]any{
			"contentType": "text",
			"content":     body,
		}
	}

	return payload
}

func buildUpdatePayload(input UpdateEventInput) map[string]any {
	payload := map[string]any{}

	if input.Subject != nil {
		payload["subject"] = *input.Subject
	}
	if input.AllDay != nil {
		payload["isAllDay"] = *input.AllDay
	}
	if input.Start != nil {
		payload["start"] = map[string]any{
			"dateTime": input.Start.Format(time.RFC3339),
			"timeZone": normalizeOptionalTimeZone(input.TimeZone),
		}
	}
	if input.End != nil {
		payload["end"] = map[string]any{
			"dateTime": input.End.Format(time.RFC3339),
			"timeZone": normalizeOptionalTimeZone(input.TimeZone),
		}
	}
	if input.Location != nil {
		payload["location"] = map[string]any{"displayName": *input.Location}
	}
	if input.Body != nil {
		payload["body"] = map[string]any{
			"contentType": "text",
			"content":     *input.Body,
		}
	}

	return payload
}

func normalizeTimeZone(value string) string {
	value = strings.TrimSpace(value)
	if value == "" {
		return "UTC"
	}
	return value
}

func normalizeOptionalTimeZone(value *string) string {
	if value == nil {
		return "UTC"
	}
	return normalizeTimeZone(*value)
}
