// vybe-test: go/http_handler_patterns/http_status_see_other
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _ = http.StatusSeeOther == 303 }
