// vybe-test: go/net_http_compile/url_parse_relative_path_only
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("/files/report.pdf")
_ = u.Path }
