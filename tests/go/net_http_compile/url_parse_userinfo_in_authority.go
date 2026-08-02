// vybe-test: go/net_http_compile/url_parse_userinfo_in_authority
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://user:pass@api.example.com/v1")
_ = u.User }
