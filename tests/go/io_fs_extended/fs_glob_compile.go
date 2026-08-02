// vybe-test: go/io_fs_extended/fs_glob_compile
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io/fs"
import "os"
func main() { _, _ = fs.Glob(os.DirFS("."), "*.go") }
