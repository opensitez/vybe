// vybe-test: go/cover_hash_heap_io/bufio_peek
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { r := bufio.NewReader(strings.NewReader("ab"))
_, _ = r.Peek(1) }
