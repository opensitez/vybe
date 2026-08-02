// vybe-test: go/cover_hash_heap_io/crc64_new
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "hash/crc64"
func main() { _ = crc64.New(crc64.MakeTable(crc64.ISO)) }
