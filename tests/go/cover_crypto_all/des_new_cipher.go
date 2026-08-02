// vybe-test: go/cover_crypto_all/des_new_cipher
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/des"
func main() { _, _ = des.NewCipher(make([]byte, 8)) }
