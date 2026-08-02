// vybe-test: go/os_exec_compile/filepath_join_two_segments
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Join("dir", "file.txt") }
