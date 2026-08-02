// vybe-test: go/hash_crc32_adler32/crc32_new_castagnoli
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs
// vybe-test-mode: compile

package main
import "hash/crc32"
func main() { h := crc32.New(crc32.MakeTable(crc32.Castagnoli))
_, _ = h.Write([]byte("x")) }
