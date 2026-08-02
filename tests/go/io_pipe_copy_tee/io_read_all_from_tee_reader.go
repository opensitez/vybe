// vybe-test: go/io_pipe_copy_tee/io_read_all_from_tee_reader
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "bytes"
import "strings"
func main() { var side bytes.Buffer
tr := io.TeeReader(strings.NewReader("all"), &side)
_, _ = io.ReadAll(tr) }
