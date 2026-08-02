// vybe-test: go/io_pipe_copy_tee/io_multi_reader_four_segments
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { mr := io.MultiReader(strings.NewReader("1"), strings.NewReader("2"), strings.NewReader("3"), strings.NewReader("4"))
_, _ = io.ReadAll(mr) }
