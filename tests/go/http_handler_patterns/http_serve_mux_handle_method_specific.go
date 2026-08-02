// vybe-test: go/http_handler_patterns/http_serve_mux_handle_method_specific
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { mux := http.NewServeMux()
mux.HandleFunc("GET /items", func(w http.ResponseWriter, r *http.Request) {}) }
