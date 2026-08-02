// vybe-test: go/io_fs_extended/io_multi_writer
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io"
import "os"
func main() { _ = io.MultiWriter(os.Stdout, os.Stderr) }
