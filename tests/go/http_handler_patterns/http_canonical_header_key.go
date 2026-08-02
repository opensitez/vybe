// vybe-test: go/http_handler_patterns/http_canonical_header_key
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _ = http.CanonicalHeaderKey("content-type") }
