// vybe-test: go/os_process_environ/os_stat_is_permission_check
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _, err := os.Stat("/etc/shadow")
_ = os.IsPermission(err) }
