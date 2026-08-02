// vybe-test: go/io_pipe_copy_tee/io_pipe_write_read_goroutine
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
func main() { pr, pw := io.Pipe()
go func() { _, _ = pw.Write([]byte("pipe"))
pw.Close() }()
_, _ = io.ReadAll(pr) }
