// vybe-test: go/os_exec_compile/filepath_ext_of_joined_path
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Ext(filepath.Join("src", "main.go")) }
