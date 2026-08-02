// vybe-test: go/crypto_cipher_modes/cipher_gcm_open_into_dst
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
gcm, _ := cipher.NewGCM(block)
nonce := make([]byte, gcm.NonceSize())
sealed := gcm.Seal(nil, nonce, []byte("msg"), nil)
_, _ = gcm.Open(sealed[:0], nonce, sealed, nil) }
