// vybe-test: go/cover_http_extra/pprof_start_cpu
// origin: languages/go/tests/go/test_cover_http_extra.rs
// vybe-test-mode: compile

package main
import "runtime/pprof"
import "os"
func main() { _ = pprof.StartCPUProfile(os.Stdout) }
