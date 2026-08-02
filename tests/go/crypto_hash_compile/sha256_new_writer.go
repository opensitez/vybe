// vybe-test: go/crypto_hash_compile/sha256_new_writer
// origin: languages/go/tests/go/test_crypto_hash_compile.rs
// vybe-test-mode: compile

package main
import "crypto/sha256"
func main() { h := sha256.New()
_, _ = h.Write([]byte("data")) }
