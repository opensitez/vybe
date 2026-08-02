// vybe-test: go/crypto_cipher_modes/aes_new_cipher_copy_key
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
func main() { key := []byte{0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15}
block, _ := aes.NewCipher(key)
key[0] = 99
plain := make([]byte, 16)
cipher := make([]byte, 16)
block.Encrypt(cipher, plain) }
