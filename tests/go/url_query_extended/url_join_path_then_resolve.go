// vybe-test: go/url_query_extended/url_join_path_then_resolve
// origin: languages/go/tests/go/test_url_query_extended.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { base, _ := url.Parse("https://ex.com/a/")
joined, _ := url.JoinPath(base.Path, "b")
ref, _ := url.Parse(joined)
_ = base.ResolveReference(ref).Path }
