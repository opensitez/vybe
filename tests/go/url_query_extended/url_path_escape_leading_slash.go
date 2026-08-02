// vybe-test: go/url_query_extended/url_path_escape_leading_slash
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { _ = url.PathEscape("/root/etc") }
