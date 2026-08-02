// vybe-test: go/url_query_extended/url_request_uri_rebuild
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.ParseRequestURI("/x?y=1")
_ = u.RequestURI() }
