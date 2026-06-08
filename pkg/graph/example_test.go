package graph

import (
	"encoding/json"
	"fmt"
	"net/http"
	"net/http/httptest"

	api "github.com/fossteams/teams-api/pkg"
	"github.com/dgrijalva/jwt-go"
)

func ExampleCalendarClient_ListMyEvents() {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		_ = json.NewEncoder(w).Encode(map[string]any{
			"value": []map[string]any{
				{
					"id":      "evt-1",
					"subject": "Design review",
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
	events, err := client.ListMyEvents()
	if err != nil {
		panic(err)
	}

	fmt.Printf("%d %s\n", len(events), events[0].Subject)
	// Output:
	// 1 Design review
}
