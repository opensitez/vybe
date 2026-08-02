// vybe-test: go/os_process_environ/os_args_copy_to_local_slice
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { dup := append([]string(nil), os.Args...)
_ = dup }
