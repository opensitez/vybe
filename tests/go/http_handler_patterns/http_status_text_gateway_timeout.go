// vybe-test: go/http_handler_patterns/http_status_text_gateway_timeout
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { _ = http.StatusText(http.StatusGatewayTimeout) }
