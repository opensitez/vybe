// vybe-test: go/io_fs_extended/fs_walk_dir_compile
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io/fs"
import "os"
func main() { _ = fs.WalkDir(os.DirFS("."), ".", func(path string, d fs.DirEntry, err error) error { return nil }) }
