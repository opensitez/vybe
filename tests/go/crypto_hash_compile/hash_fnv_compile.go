// vybe-test: go/crypto_hash_compile/hash_fnv_compile
// origin: languages/go/tests/go/test_crypto_hash_compile.rs
// vybe-test-mode: compile

package main
import "hash/fnv"
func main() { h := fnv.New32a()
_, _ = h.Write([]byte("a")) }
