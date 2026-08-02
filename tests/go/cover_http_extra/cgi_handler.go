// vybe-test: go/cover_http_extra/cgi_handler
// origin: languages/go/tests/go/test_cover_http_extra.rs
// vybe-test-mode: compile

package main
import "net/http/cgi"
func main() { _ = cgi.Handler{} }
