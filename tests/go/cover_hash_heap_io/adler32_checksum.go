// vybe-test: go/cover_hash_heap_io/adler32_checksum
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "hash/adler32"
func main() { _ = adler32.Checksum([]byte("go")) }
