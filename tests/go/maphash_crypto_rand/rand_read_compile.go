// vybe-test: go/maphash_crypto_rand/rand_read_compile
// origin: languages/go/tests/go/test_maphash_crypto_rand.rs
// vybe-test-mode: compile

package main
import "crypto/rand"
func main() { b := make([]byte, 8)
_, _ = rand.Read(b) }
