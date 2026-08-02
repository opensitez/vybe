// vybe-test: go/bufio_io/reader_reset_new_input
// origin: languages/go/tests/go/test_bufio_io.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { r := bufio.NewReader(strings.NewReader("old"))
r.Reset(strings.NewReader("new")) }
