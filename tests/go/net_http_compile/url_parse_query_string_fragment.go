// vybe-test: go/net_http_compile/url_parse_query_string_fragment
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://host/search?q=go&lang=en#results")
_ = u.RawQuery
_ = u.Fragment }
