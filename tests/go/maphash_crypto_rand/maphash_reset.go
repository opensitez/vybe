// vybe-test: go/maphash_crypto_rand/maphash_reset
// origin: languages/go/tests/go/test_maphash_crypto_rand.rs
// vybe-test-mode: compile

package main
import "hash/maphash"
func main() { var h maphash.Hash
h.Reset() }
