// vybe-test: go/cover_http_extra/httptest_response_recorder
// origin: languages/go/tests/go/test_cover_http_extra.rs
// vybe-test-mode: compile

package main
import "net/http/httptest"
func main() { _ = httptest.NewRecorder() }
