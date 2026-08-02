// vybe-test: go/cover_hash_heap_io/compress_lzw_writer
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "compress/lzw"
import "bytes"
func main() { _ = lzw.NewWriter(bytes.NewBuffer(nil), lzw.LSB, 8) }
