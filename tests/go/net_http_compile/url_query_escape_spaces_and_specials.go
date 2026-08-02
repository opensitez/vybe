// vybe-test: go/net_http_compile/url_query_escape_spaces_and_specials
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { _ = url.QueryEscape("a b&c=d") }
