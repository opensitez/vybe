// vybe-test: go/crypto_hash_compile/sha1_sum_compile
// origin: languages/go/tests/go/test_crypto_hash_compile.rs
// vybe-test-mode: compile

package main
import "crypto/sha1"
func main() { _ = sha1.Sum([]byte("x")) }
