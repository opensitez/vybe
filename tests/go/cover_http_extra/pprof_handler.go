// vybe-test: go/cover_http_extra/pprof_handler
// origin: languages/go/tests/go/test_cover_http_extra.rs
// vybe-test-mode: compile

package main
import "net/http/pprof"
func main() { _ = pprof.Handler("goroutine") }
