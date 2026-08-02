// vybe-test: go/cover_hash_heap_io/compress_bzip2_reader
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "compress/bzip2"
import "bytes"
func main() { _ = bzip2.NewReader(bytes.NewReader(nil)) }
