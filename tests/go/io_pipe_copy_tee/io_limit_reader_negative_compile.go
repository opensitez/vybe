// vybe-test: go/io_pipe_copy_tee/io_limit_reader_negative_compile
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { _ = io.LimitReader(strings.NewReader("a"), -1) }
