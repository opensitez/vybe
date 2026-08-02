// vybe-test: go/io_fs_extended/io_pipe_compile
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io"
func main() { r, w := io.Pipe()
_ = r
_ = w }
