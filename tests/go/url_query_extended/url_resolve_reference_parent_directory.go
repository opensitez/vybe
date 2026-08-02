// vybe-test: go/url_query_extended/url_resolve_reference_parent_directory
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { base, _ := url.Parse("https://ex.com/a/b/c")
ref, _ := url.Parse("../d")
_ = base.ResolveReference(ref).Path }
