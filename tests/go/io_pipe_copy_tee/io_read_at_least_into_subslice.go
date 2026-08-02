// vybe-test: go/io_pipe_copy_tee/io_read_at_least_into_subslice
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { buf := make([]byte, 4)
_, _ = io.ReadAtLeast(strings.NewReader("go"), buf[:2], 2) }
