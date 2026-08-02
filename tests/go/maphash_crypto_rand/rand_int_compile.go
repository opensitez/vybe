// vybe-test: go/maphash_crypto_rand/rand_int_compile
// origin: languages/go/tests/go/test_maphash_crypto_rand.rs
// vybe-test-mode: compile

package main
import "crypto/rand"
import "math/big"
func main() { _, _ = rand.Int(rand.Reader, big.NewInt(10)) }
