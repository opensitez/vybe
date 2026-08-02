// vybe-test: go/cover_http_extra/httptest_new_request
// origin: languages/go/tests/go/test_cover_http_extra.rs
// vybe-test-mode: compile

package main
import "net/http/httptest"
func main() { _ = httptest.NewRequest("GET", "/", nil) }
