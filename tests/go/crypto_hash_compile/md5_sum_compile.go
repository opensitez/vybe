// vybe-test: go/crypto_hash_compile/md5_sum_compile
// origin: languages/go/tests/go/test_crypto_hash_compile.rs
// vybe-test-mode: compile

package main
import "crypto/md5"
func main() { _ = md5.Sum([]byte("x")) }
