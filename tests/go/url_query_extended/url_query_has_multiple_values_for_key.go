// vybe-test: go/url_query_extended/url_query_has_multiple_values_for_key
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { u, _ := url.Parse("https://h/?tag=a&tag=b")
_ = u.Query()["tag"] }
