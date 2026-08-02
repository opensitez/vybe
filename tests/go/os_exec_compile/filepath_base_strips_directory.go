// vybe-test: go/os_exec_compile/filepath_base_strips_directory
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Base("/var/log/app.log") }
