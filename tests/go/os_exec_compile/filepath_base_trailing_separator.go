// vybe-test: go/os_exec_compile/filepath_base_trailing_separator
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Base("/tmp/build/") }
