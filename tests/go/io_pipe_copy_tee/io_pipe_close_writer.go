// vybe-test: go/io_pipe_copy_tee/io_pipe_close_writer
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
func main() { pr, pw := io.Pipe()
_ = pw.Close()
_, _ = pr.Read(make([]byte, 1)) }
