package graph

import (
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
	"time"

	api "github.com/saimon-moore/teams-api/pkg"
	"github.com/dgrijalva/jwt-go"
)

func testCalendarClient(serverURL string) *CalendarClient {
	token := &api.TeamsToken{
		Inner: &jwt.Token{Raw: "graph-token"},
		Type:  api.TokenBearer,
	}

	return NewCalendarClientWithBaseURL(serverURL, http.DefaultClient, token)
}

func TestListCalendars(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v1.0/me/calendars" {
			t.Fatalf("expected calendars path, got %s", r.URL.Path)
		}
		if got := r.Header.Get("Authorization"); got != "Bearer graph-token" {
			t.Fatalf("expected bearer auth header, got %q", got)
		}

		_ = json.NewEncoder(w).Encode(map[string]any{
			"value": []map[string]any{
				{
					"id":                "cal-1",
					"name":              "Primary",
					"canEdit":           true,
					"isDefaultCalendar": true,
					"owner": map[string]any{
						"name":    "Dev User",
						"address": "dev@example.com",
					},
				},
			},
		})
	}))
	defer server.Close()

	calendars, err := testCalendarClient(server.URL).ListCalendars()
	if err != nil {
		t.Fatalf("expected calendars response, got error: %v", err)
	}

	if len(calendars) != 1 {
		t.Fatalf("expected 1 calendar, got %d", len(calendars))
	}
	if calendars[0].ID != "cal-1" {
		t.Fatalf("expected calendar id cal-1, got %s", calendars[0].ID)
	}
	if calendars[0].Owner.Address != "dev@example.com" {
		t.Fatalf("expected owner address dev@example.com, got %s", calendars[0].Owner.Address)
	}
}

func TestListEventsUsesDefaultCalendarView(t *testing.T) {
	start := time.Date(2026, time.June, 8, 9, 0, 0, 0, time.UTC)
	end := start.Add(24 * time.Hour)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v1.0/me/calendarView" {
			t.Fatalf("expected default calendar view path, got %s", r.URL.Path)
		}
		if got := r.URL.Query().Get("startDateTime"); got != start.Format(time.RFC3339) {
			t.Fatalf("expected startDateTime %s, got %s", start.Format(time.RFC3339), got)
		}
		if got := r.URL.Query().Get("endDateTime"); got != end.Format(time.RFC3339) {
			t.Fatalf("expected endDateTime %s, got %s", end.Format(time.RFC3339), got)
		}
		if got := r.URL.Query().Get("$top"); got != "25" {
			t.Fatalf("expected $top 25, got %s", got)
		}
		if got := r.Header.Get("Prefer"); got != `outlook.timezone="Europe/Zurich"` {
			t.Fatalf("expected timezone preference header, got %q", got)
		}

		_ = json.NewEncoder(w).Encode(map[string]any{
			"value": []map[string]any{
				{
					"id":        "evt-1",
					"subject":   "Design review",
					"isAllDay":  false,
					"webLink":   "https://example.test/event",
					"bodyPreview": "Short body",
					"start": map[string]any{
						"dateTime": start.Format(time.RFC3339),
						"timeZone": "UTC",
					},
					"end": map[string]any{
						"dateTime": end.Format(time.RFC3339),
						"timeZone": "UTC",
					},
					"location": map[string]any{
						"displayName": "Room 1",
					},
				},
			},
		})
	}))
	defer server.Close()

	events, err := testCalendarClient(server.URL).ListEvents(ListEventsOptions{
		Start:    start,
		End:      end,
		Limit:    25,
		TimeZone: "Europe/Zurich",
	})
	if err != nil {
		t.Fatalf("expected events response, got error: %v", err)
	}

	if len(events) != 1 {
		t.Fatalf("expected 1 event, got %d", len(events))
	}
	if events[0].Location.DisplayName != "Room 1" {
		t.Fatalf("expected location Room 1, got %s", events[0].Location.DisplayName)
	}
	if events[0].BodyPreview != "Short body" {
		t.Fatalf("expected body preview to round-trip, got %q", events[0].BodyPreview)
	}
}

func TestListEventsUsesExplicitCalendarID(t *testing.T) {
	start := time.Date(2026, time.June, 8, 9, 0, 0, 0, time.UTC)
	end := start.Add(2 * time.Hour)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v1.0/me/calendars/cal-42/calendarView" {
			t.Fatalf("expected explicit calendar path, got %s", r.URL.Path)
		}
		_ = json.NewEncoder(w).Encode(map[string]any{"value": []map[string]any{}})
	}))
	defer server.Close()

	_, err := testCalendarClient(server.URL).ListEvents(ListEventsOptions{
		CalendarID: "cal-42",
		Start:      start,
		End:        end,
	})
	if err != nil {
		t.Fatalf("expected events response, got error: %v", err)
	}
}

func TestCreateEventUsesDefaultCalendarEndpoint(t *testing.T) {
	start := time.Date(2026, time.June, 8, 9, 0, 0, 0, time.UTC)
	end := start.Add(time.Hour)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			t.Fatalf("expected POST, got %s", r.Method)
		}
		if r.URL.Path != "/v1.0/me/events" {
			t.Fatalf("expected default create path, got %s", r.URL.Path)
		}

		body, err := io.ReadAll(r.Body)
		if err != nil {
			t.Fatalf("unable to read request body: %v", err)
		}

		payload := string(body)
		for _, needle := range []string{
			`"subject":"Design review"`,
			`"isAllDay":true`,
			`"timeZone":"Europe/Zurich"`,
			`"displayName":"Room 1"`,
			`"content":"Agenda"`,
		} {
			if !strings.Contains(payload, needle) {
				t.Fatalf("expected request body to contain %s, got %s", needle, payload)
			}
		}

		w.WriteHeader(http.StatusCreated)
		_ = json.NewEncoder(w).Encode(map[string]any{
			"id":      "evt-1",
			"subject": "Design review",
		})
	}))
	defer server.Close()

	event, err := testCalendarClient(server.URL).CreateEvent(CreateEventInput{
		Subject:  "Design review",
		Start:    start,
		End:      end,
		TimeZone: "Europe/Zurich",
		Location: "Room 1",
		Body:     "Agenda",
		AllDay:   true,
	})
	if err != nil {
		t.Fatalf("expected create response, got error: %v", err)
	}
	if event.ID != "evt-1" {
		t.Fatalf("expected created event id evt-1, got %s", event.ID)
	}
}

func TestUpdateEventUsesExplicitCalendarEndpointAndPartialPayload(t *testing.T) {
	subject := "Updated title"

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPatch {
			t.Fatalf("expected PATCH, got %s", r.Method)
		}
		if r.URL.Path != "/v1.0/me/calendars/cal-42/events/evt-9" {
			t.Fatalf("expected explicit update path, got %s", r.URL.Path)
		}

		body, err := io.ReadAll(r.Body)
		if err != nil {
			t.Fatalf("unable to read request body: %v", err)
		}

		payload := string(body)
		if !strings.Contains(payload, `"subject":"Updated title"`) {
			t.Fatalf("expected partial update payload, got %s", payload)
		}
		if strings.Contains(payload, `"start"`) {
			t.Fatalf("did not expect start field in partial payload, got %s", payload)
		}

		_ = json.NewEncoder(w).Encode(map[string]any{
			"id":      "evt-9",
			"subject": subject,
		})
	}))
	defer server.Close()

	event, err := testCalendarClient(server.URL).UpdateEvent("cal-42", "evt-9", UpdateEventInput{
		Subject: &subject,
	})
	if err != nil {
		t.Fatalf("expected update response, got error: %v", err)
	}
	if event.Subject != subject {
		t.Fatalf("expected updated subject %q, got %q", subject, event.Subject)
	}
}

func TestDeleteEventUsesDefaultCalendarEndpoint(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodDelete {
			t.Fatalf("expected DELETE, got %s", r.Method)
		}
		if r.URL.Path != "/v1.0/me/events/evt-7" {
			t.Fatalf("expected default delete path, got %s", r.URL.Path)
		}
		w.WriteHeader(http.StatusNoContent)
	}))
	defer server.Close()

	if err := testCalendarClient(server.URL).DeleteEvent("", "evt-7"); err != nil {
		t.Fatalf("expected delete response, got error: %v", err)
	}
}

func TestListEventsRequiresRange(t *testing.T) {
	_, err := testCalendarClient("https://graph.microsoft.com").ListEvents(ListEventsOptions{})
	if err == nil {
		t.Fatal("expected range validation error")
	}
	if !strings.Contains(err.Error(), "start and end") {
		t.Fatalf("expected range validation error, got %v", err)
	}
}

func TestGraphClientReturnsUsefulStatusErrors(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		http.Error(w, "graph said no", http.StatusForbidden)
	}))
	defer server.Close()

	_, err := testCalendarClient(server.URL).ListCalendars()
	if err == nil {
		t.Fatal("expected status error")
	}
	if !strings.Contains(err.Error(), "403") || !strings.Contains(err.Error(), "graph said no") {
		t.Fatalf("expected detailed status error, got %v", err)
	}
}
