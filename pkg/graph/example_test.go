package graph

import (
	"encoding/json"
	"fmt"
	"net/http"
	"net/http/httptest"
	"time"

	api "github.com/saimon-moore/teams-api/pkg"
	"github.com/dgrijalva/jwt-go"
)

func ExampleCalendarClient_ListEvents() {
	start := time.Date(2026, time.June, 8, 9, 0, 0, 0, time.UTC)
	end := start.Add(24 * time.Hour)

	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_ = json.NewEncoder(w).Encode(map[string]any{
			"value": []map[string]any{
				{
					"id":      "evt-1",
					"subject": "Design review",
					"start": map[string]any{
						"dateTime": start.Format(time.RFC3339),
						"timeZone": "UTC",
					},
					"end": map[string]any{
						"dateTime": end.Format(time.RFC3339),
						"timeZone": "UTC",
					},
				},
			},
		})
	}))
	defer server.Close()

	token := &api.TeamsToken{
		Inner: &jwt.Token{Raw: "graph-token"},
		Type:  api.TokenBearer,
	}

	client := NewCalendarClientWithBaseURL(server.URL, http.DefaultClient, token)
	events, err := client.ListEvents(ListEventsOptions{
		Start: start,
		End:   end,
	})
	if err != nil {
		panic(err)
	}

	fmt.Printf("%d %s\n", len(events), events[0].Subject)
	// Output:
	// 1 Design review
}
