// vybe-test: go/io_pipe_copy_tee/io_tee_reader_to_discard
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { tr := io.TeeReader(strings.NewReader("d"), io.Discard)
_, _ = io.ReadAll(tr) }
