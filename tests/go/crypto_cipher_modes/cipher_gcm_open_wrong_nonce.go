// vybe-test: go/crypto_cipher_modes/cipher_gcm_open_wrong_nonce
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
gcm, _ := cipher.NewGCM(block)
nonce := make([]byte, gcm.NonceSize())
sealed := gcm.Seal(nil, nonce, []byte("x"), nil)
badNonce := make([]byte, gcm.NonceSize())
badNonce[0] = 1
_, err := gcm.Open(nil, badNonce, sealed, nil)
_ = err }
