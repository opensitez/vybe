// vybe-test: go/os_exec_compile/getenv_passed_to_local_helper
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "os"
func pick(key string) string { return os.Getenv(key) }
func main() { _ = pick("PATH") }
