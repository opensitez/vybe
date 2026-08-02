// vybe-test: go/url_query_extended/url_escaped_path_decodes_percent
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://h/p%2Fq")
_ = u.EscapedPath() }
