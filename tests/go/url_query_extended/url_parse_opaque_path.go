// vybe-test: go/url_query_extended/url_parse_opaque_path
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("mailto:user@example.com")
_ = u.Opaque }
