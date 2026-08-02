// vybe-test: go/url_query_extended/url_parse_request_uri_with_userinfo
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.ParseRequestURI("https://u:p@host/res")
_ = u.User }
