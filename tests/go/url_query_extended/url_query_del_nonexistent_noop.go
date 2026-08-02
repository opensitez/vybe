// vybe-test: go/url_query_extended/url_query_del_nonexistent_noop
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { q := url.Values{"a": []string{"1"}}
q.Del("z")
_ = len(q) }
