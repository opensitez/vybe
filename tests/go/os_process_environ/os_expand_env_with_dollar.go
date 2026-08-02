// vybe-test: go/os_process_environ/os_expand_env_with_dollar
// origin: languages/go/tests/go/test_os_process_environ.rs
// vybe-test-mode: compile

package main
import "os"
func main() { _ = os.ExpandEnv("${HOME}/bin") }
