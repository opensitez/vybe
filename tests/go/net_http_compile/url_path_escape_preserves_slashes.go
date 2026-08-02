// vybe-test: go/net_http_compile/url_path_escape_preserves_slashes
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/url"
func main() { _ = url.PathEscape("dir/file name.txt") }
