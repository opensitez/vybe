// vybe-test: go/url_query_extended/url_join_path_empty_elements_skipped
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { _, _ = url.JoinPath("a", "", "b") }
