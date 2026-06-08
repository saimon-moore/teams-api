package api

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/dgrijalva/jwt-go"
)

func TestGetGraphTokenFromEnv(t *testing.T) {
	t.Setenv("MS_TEAMS_GRAPH_TOKEN", testJWT(t))

	token, err := GetGraphToken()
	if err != nil {
		t.Fatalf("expected graph token from env, got error: %v", err)
	}

	if token == nil || token.Inner == nil {
		t.Fatal("expected parsed graph token")
	}
	if token.Inner.Raw != os.Getenv("MS_TEAMS_GRAPH_TOKEN") {
		t.Fatalf("expected raw env token to round-trip, got %q", token.Inner.Raw)
	}
	if token.Type != TokenBearer {
		t.Fatalf("expected bearer token type, got %s", token.Type)
	}
}

func TestGetGraphTokenFromFile(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("HOME", dir)
	t.Setenv("MS_TEAMS_GRAPH_TOKEN", "")

	tokenDir := filepath.Join(dir, ".config", "fossteams")
	if err := os.MkdirAll(tokenDir, 0o755); err != nil {
		t.Fatalf("unable to create token dir: %v", err)
	}

	want := testJWT(t)
	if err := os.WriteFile(filepath.Join(tokenDir, "token-graph.jwt"), []byte(want), 0o600); err != nil {
		t.Fatalf("unable to write graph token file: %v", err)
	}

	token, err := GetGraphToken()
	if err != nil {
		t.Fatalf("expected graph token from file, got error: %v", err)
	}

	if token == nil || token.Inner == nil {
		t.Fatal("expected parsed graph token")
	}
	if token.Inner.Raw != want {
		t.Fatalf("expected raw file token to round-trip, got %q", token.Inner.Raw)
	}
	if token.Type != TokenBearer {
		t.Fatalf("expected bearer token type, got %s", token.Type)
	}
}

func testJWT(t *testing.T) string {
	t.Helper()

	token := jwt.NewWithClaims(jwt.SigningMethodHS256, jwt.MapClaims{
		"aud": "https://graph.microsoft.com",
		"sub": "graph-user",
	})

	raw, err := token.SignedString([]byte("test-secret"))
	if err != nil {
		t.Fatalf("unable to sign JWT: %v", err)
	}

	return raw
}
