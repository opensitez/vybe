// vybe-test: go/crypto_cipher_modes/cipher_gcm_seal_empty_plaintext
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
gcm, _ := cipher.NewGCM(block)
nonce := make([]byte, gcm.NonceSize())
_ = gcm.Seal(nil, nonce, []byte{}, nil) }
