// vybe-test: go/http_handler_patterns/http_not_found_helper
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { http.NotFound(w, r) }) }
