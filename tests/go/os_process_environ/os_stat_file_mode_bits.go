// vybe-test: go/os_process_environ/os_stat_file_mode_bits
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { fi, err := os.Stat(".")
if err == nil { _ = fi.Mode() } }
