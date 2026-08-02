// vybe-test: go/url_query_extended/url_redacted_omits_credentials
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://secret:pass@host/path")
_ = u.Redacted() }
