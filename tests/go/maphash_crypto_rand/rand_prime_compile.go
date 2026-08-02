// vybe-test: go/maphash_crypto_rand/rand_prime_compile
// origin: languages/go/tests/go/test_maphash_crypto_rand.rs
// vybe-test-mode: compile

package main
import "crypto/rand"
import "math/big"
func main() { _, _ = rand.Prime(rand.Reader, 16) }
