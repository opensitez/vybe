// vybe-test: go/cover_hash_heap_io/io_copybuffer
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
import "bytes"
func main() { _, _ = io.CopyBuffer(bytes.NewBuffer(nil), strings.NewReader("a"), make([]byte, 8)) }
