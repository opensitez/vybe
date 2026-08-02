// vybe-test: go/http_handler_patterns/http_request_host_field
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://api.example.com/v1", nil)
_ = req.Host }
