// vybe-test: go/net_http_compile/http_post_with_content_type_header
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
import "strings"
func main() { _, _ = http.Post("https://example.com/form", "application/x-www-form-urlencoded", strings.NewReader("a=1")) }
