// vybe-test: go/crypto_cipher_modes/cipher_new_gcm_seal
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
import "crypto/rand"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
gcm, _ := cipher.NewGCM(block)
nonce := make([]byte, gcm.NonceSize())
plain := []byte("secret")
_ = gcm.Seal(nil, nonce, plain, nil) }
