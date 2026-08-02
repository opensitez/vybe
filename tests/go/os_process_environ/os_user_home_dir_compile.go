// vybe-test: go/os_process_environ/os_user_home_dir_compile
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { home, err := os.UserHomeDir()
_ = home
_ = err }
