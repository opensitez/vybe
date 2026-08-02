// vybe-test: go/http_handler_patterns/http_status_temporary_redirect
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _ = http.StatusTemporaryRedirect == 307 }
