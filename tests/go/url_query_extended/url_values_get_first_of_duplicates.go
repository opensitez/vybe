// vybe-test: go/url_query_extended/url_values_get_first_of_duplicates
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { v := url.Values{"k": []string{"first", "second"}}
_ = v.Get("k") }
