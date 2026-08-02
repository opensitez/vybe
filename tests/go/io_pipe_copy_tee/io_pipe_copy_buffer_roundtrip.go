// vybe-test: go/io_pipe_copy_tee/io_pipe_copy_buffer_roundtrip
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { pr, pw := io.Pipe()
go func() { buf := make([]byte, 8)
_, _ = io.CopyBuffer(pw, strings.NewReader("buf"), buf)
pw.Close() }()
_, _ = io.ReadAll(pr) }
