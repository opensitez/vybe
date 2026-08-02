// vybe-test: go/http_handler_patterns/http_status_permanent_redirect
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _ = http.StatusPermanentRedirect == 308 }
