// vybe-test: go/os_process_environ/os_same_file_stat_compare
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { a, e1 := os.Stat(".")
b, e2 := os.Lstat(".")
if e1 == nil && e2 == nil { _, _ = os.SameFile(a, b) } }
