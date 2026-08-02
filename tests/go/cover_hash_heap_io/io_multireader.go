// vybe-test: go/cover_hash_heap_io/io_multireader
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { _ = io.MultiReader(strings.NewReader("a"), strings.NewReader("b")) }
