// vybe-test: go/http_handler_patterns/http_request_proto_major_minor
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
func main() { req, _ := http.NewRequest(http.MethodGet, "https://ex.com", nil)
_ = req.ProtoMajor
_ = req.ProtoMinor }
