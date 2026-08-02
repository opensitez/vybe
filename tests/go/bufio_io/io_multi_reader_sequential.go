// vybe-test: go/bufio_io/io_multi_reader_sequential
// origin: languages/go/tests/go/test_bufio_io.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { mr := io.MultiReader(strings.NewReader("ab"), strings.NewReader("cd"))
_, _ = io.ReadAll(mr) }
