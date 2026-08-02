// vybe-test: go/bufio_io/io_pipe_roundtrip
// origin: languages/go/tests/go/test_bufio_io.rs
// vybe-test-mode: compile

package main
import "io"
func main() { pr, pw := io.Pipe()
go func() { _, _ = pw.Write([]byte("p"))
pw.Close() }()
_, _ = io.ReadAll(pr) }
