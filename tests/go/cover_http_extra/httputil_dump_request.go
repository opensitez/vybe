// vybe-test: go/cover_http_extra/httputil_dump_request
// origin: languages/go/tests/go/test_cover_http_extra.rs
// vybe-test-mode: compile

package main
import "net/http/httputil"
func main() { _, _ = httputil.DumpRequest(nil, false) }
