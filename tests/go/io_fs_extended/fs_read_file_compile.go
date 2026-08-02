// vybe-test: go/io_fs_extended/fs_read_file_compile
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io/fs"
import "os"
func main() { _, _ = fs.ReadFile(os.DirFS("."), "go.mod") }
