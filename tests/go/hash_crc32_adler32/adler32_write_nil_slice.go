// vybe-test: go/hash_crc32_adler32/adler32_write_nil_slice
// origin: languages/go/tests/go/test_hash_crc32_adler32.rs
// vybe-test-mode: compile

package main
import "hash/adler32"
func main() { h := adler32.New()
_, _ = h.Write(nil) }
