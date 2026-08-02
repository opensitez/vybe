// vybe-test: go/net_http_compile/url_parse_request_uri_no_fragment
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.ParseRequestURI("/index.html?tab=1")
_ = u.Path }
