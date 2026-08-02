// vybe-test: go/url_query_extended/url_resolve_reference_query_merge
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { base, _ := url.Parse("https://ex.com/a?x=1")
ref, _ := url.Parse("?y=2")
_ = base.ResolveReference(ref).RawQuery }
