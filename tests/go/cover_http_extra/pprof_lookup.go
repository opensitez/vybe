// vybe-test: go/cover_http_extra/pprof_lookup
// origin: languages/go/tests/go/test_cover_http_extra.rs
// vybe-test-mode: compile

package main
import "runtime/pprof"
func main() { _, _ = pprof.Lookup("goroutine") }
