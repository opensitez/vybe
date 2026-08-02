// vybe-test: go/http_handler_patterns/http_serve_mux_handle_pattern
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { mux := http.NewServeMux()
mux.Handle("/api/", http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {})) }
