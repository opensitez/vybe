// vybe-test: go/crypto_hash_compile/hash_crc32_compile
// origin: languages/go/tests/go/test_crypto_hash_compile.rs
// vybe-test-mode: compile

package main
import "hash/crc32"
func main() { _ = crc32.ChecksumIEEE([]byte("go")) }
