// vybe-test: go/http_handler_patterns/http_default_serve_mux_handle_func
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.HandleFunc("/ping", func(w http.ResponseWriter, r *http.Request) {}) }
