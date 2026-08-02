// vybe-test: go/url_query_extended/url_parse_request_uri_with_percent_encoded
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.ParseRequestURI("/files/hello%20world.txt")
_ = u.Path }
