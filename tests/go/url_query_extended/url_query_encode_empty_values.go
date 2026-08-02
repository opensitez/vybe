// vybe-test: go/url_query_extended/url_query_encode_empty_values
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { q := url.Values{"empty": []string{""}}
_ = q.Encode() }
