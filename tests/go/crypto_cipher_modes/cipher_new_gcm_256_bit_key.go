// vybe-test: go/crypto_cipher_modes/cipher_new_gcm_256_bit_key
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 32))
gcm, _ := cipher.NewGCM(block)
_ = gcm.NonceSize() }
