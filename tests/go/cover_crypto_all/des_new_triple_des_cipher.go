// vybe-test: go/cover_crypto_all/des_new_triple_des_cipher
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/des"
func main() { _, _ = des.NewTripleDESCipher(make([]byte, 24)) }
