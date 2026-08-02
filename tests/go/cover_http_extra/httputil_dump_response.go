// vybe-test: go/cover_http_extra/httputil_dump_response
// origin: languages/go/tests/go/test_cover_http_extra.rs
// vybe-test-mode: compile

package main
import "net/http/httputil"
func main() { _, _ = httputil.DumpResponse(nil, false) }
