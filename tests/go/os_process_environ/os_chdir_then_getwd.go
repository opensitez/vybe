// vybe-test: go/os_process_environ/os_chdir_then_getwd
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { wd, _ := os.Getwd()
defer os.Chdir(wd)
_ = os.Chdir(".") }
