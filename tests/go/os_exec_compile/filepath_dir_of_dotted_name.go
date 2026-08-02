// vybe-test: go/os_exec_compile/filepath_dir_of_dotted_name
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Dir("/opt/bin/tool.exe") }
