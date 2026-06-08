package graph

import (
	"encoding/json"
	"net/http"
	"net/http/httptest"
	"testing"

	api "github.com/fossteams/teams-api/pkg"
	"github.com/dgrijalva/jwt-go"
)

func TestListMyEvents(t *testing.T) {
	token := &api.TeamsToken{
		Inner: &jwt.Token{Raw: "graph-token"},
		Type:  api.TokenBearer,
	}

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/v1.0/me/calendar/events" {
			t.Fatalf("expected Graph events path, got %s", r.URL.Path)
		}
		if got := r.Header.Get("Authorization"); got != "Bearer graph-token" {
			t.Fatalf("expected bearer auth header, got %q", got)
		}

		w.Header().Set("Content-Type", "application/json")
		if err := json.NewEncoder(w).Encode(map[string]any{
			"value": []map[string]any{
				{
					"id":      "evt-1",
					"subject": "Test event",
				},
			},
		}); err != nil {
			t.Fatalf("unable to encode response: %v", err)
		}
	}))
	defer server.Close()

	client := NewCalendarClientWithBaseURL(server.URL, http.DefaultClient, token)
	events, err := client.ListMyEvents()
	if err != nil {
		t.Fatalf("expected events response, got error: %v", err)
	}

	if len(events) != 1 {
		t.Fatalf("expected 1 event, got %d", len(events))
	}
	if events[0].ID != "evt-1" {
		t.Fatalf("expected event id evt-1, got %s", events[0].ID)
	}
	if events[0].Subject != "Test event" {
		t.Fatalf("expected event subject Test event, got %s", events[0].Subject)
	}
}
