// vybe-test: go/http_handler_patterns/http_serve_mux_serve_http_dispatch
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { mux := http.NewServeMux()
mux.HandleFunc("/x", func(w http.ResponseWriter, r *http.Request) {})
mux.ServeHTTP(nil, nil) }
