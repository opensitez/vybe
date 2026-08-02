// vybe-test: go/http_handler_patterns/http_max_bytes_reader_limit
// origin: languages/go/tests/go/test_http_handler_patterns.rs
// vybe-test-mode: compile

package main
import "net/http"
import "strings"
func main() { req, _ := http.NewRequest(http.MethodPost, "https://ex.com", strings.NewReader("body"))
_ = http.MaxBytesReader(nil, req.Body, 1024) }
