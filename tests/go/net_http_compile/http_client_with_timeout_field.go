// vybe-test: go/net_http_compile/http_client_with_timeout_field
// origin: languages/go/tests/go/test_net_http_compile.rs
// vybe-test-mode: compile

package main
import "net/http"
import "time"
func main() { c := http.Client{Timeout: 5 * time.Second}
_, _ = c.Get("https://slow.example/status") }
