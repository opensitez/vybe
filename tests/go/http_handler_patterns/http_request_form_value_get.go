// vybe-test: go/http_handler_patterns/http_request_form_value_get
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodPost, "https://ex.com", nil)
_ = req.FormValue("field") }
