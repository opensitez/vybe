// vybe-test: go/cover_http_extra/httptrace_context_client_trace
// origin: languages/go/tests/go/test_cover_http_extra.rs
// vybe-test-mode: compile

package main
import "net/http/httptrace"
func main() { _ = httptrace.WithClientTrace(nil, &httptrace.ClientTrace{}) }
