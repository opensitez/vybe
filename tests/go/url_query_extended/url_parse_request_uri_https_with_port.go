// vybe-test: go/url_query_extended/url_parse_request_uri_https_with_port
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.ParseRequestURI("https://localhost:8443/status")
_ = u.Host }
