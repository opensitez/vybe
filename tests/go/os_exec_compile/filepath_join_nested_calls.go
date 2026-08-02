// vybe-test: go/os_exec_compile/filepath_join_nested_calls
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Join(filepath.Join("root", "sub"), "leaf") }
