// vybe-test: go/os_exec_compile/filepath_ext_returns_suffix
// origin: languages/go/tests/go/test_os_exec_compile.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Ext("archive.tar.gz") }
