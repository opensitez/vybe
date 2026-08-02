// vybe-test: go/crypto_cipher_modes/cipher_new_gcm_open
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
gcm, _ := cipher.NewGCM(block)
nonce := make([]byte, gcm.NonceSize())
ct := gcm.Seal(nil, nonce, []byte("data"), nil)
_, _ = gcm.Open(nil, nonce, ct, nil) }
