// vybe-test: go/net_http_compile/url_parse_then_set_query_via_values
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://host/")
q := url.Values{}
q.Set("page", "2")
u.RawQuery = q.Encode()
_ = u.String() }
