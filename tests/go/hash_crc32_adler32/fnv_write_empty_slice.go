// vybe-test: go/hash_crc32_adler32/fnv_write_empty_slice
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs
// vybe-test-mode: compile

package main
import "hash/fnv"
func main() { h := fnv.New32()
_, _ = h.Write([]byte{}) }
