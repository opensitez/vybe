// vybe-test: go/crypto_hash_compile/sha256_size_constant
// origin: languages/go/tests/go/test_crypto_hash_compile.rs
// vybe-test-mode: compile

package main
import "crypto/sha256"
func main() { _ = sha256.Size }
