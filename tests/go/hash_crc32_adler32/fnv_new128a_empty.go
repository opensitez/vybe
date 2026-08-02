// vybe-test: go/hash_crc32_adler32/fnv_new128a_empty
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs
// vybe-test-mode: compile

package main
import "hash/fnv"
func main() { _ = fnv.New128a().Sum(nil) }
