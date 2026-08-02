// vybe-test: go/url_query_extended/url_resolve_reference_fragment_from_ref
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { base, _ := url.Parse("https://ex.com/page")
ref, _ := url.Parse("#section")
_ = base.ResolveReference(ref).Fragment }
