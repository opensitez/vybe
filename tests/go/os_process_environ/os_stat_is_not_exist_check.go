// vybe-test: go/os_process_environ/os_stat_is_not_exist_check
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _, err := os.Stat("/no/such/vybe/path")
_ = os.IsNotExist(err) }
