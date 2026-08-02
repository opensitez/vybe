// vybe-test: go/http_handler_patterns/http_default_serve_mux_handle
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.Handle("/static/", http.FileServer(http.Dir("."))) }
