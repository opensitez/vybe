// vybe-test: go/net_http_compile/url_parse_http_with_port
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("http://localhost:8080/api")
_ = u.Host }
