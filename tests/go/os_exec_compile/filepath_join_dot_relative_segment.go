// vybe-test: go/os_exec_compile/filepath_join_dot_relative_segment
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Join(".", "config", "app.toml") }
