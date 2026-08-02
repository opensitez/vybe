// vybe-test: go/net_http_compile/url_parse_https_absolute
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://example.com/path")
_ = u.Scheme }
