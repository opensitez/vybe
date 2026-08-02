// vybe-test: go/url_query_extended/url_path_unescape_plus_not_space
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { _, _ = url.PathUnescape("100%25") }
