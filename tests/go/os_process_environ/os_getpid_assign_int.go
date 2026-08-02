// vybe-test: go/os_process_environ/os_getpid_assign_int
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { pid := os.Getpid()
_ = pid }
