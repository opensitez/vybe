// vybe-test: go/crypto_cipher_modes/cipher_gcm_seal_dst_preallocated
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
gcm, _ := cipher.NewGCM(block)
nonce := make([]byte, gcm.NonceSize())
dst := make([]byte, 0, 64)
_ = gcm.Seal(dst, nonce, []byte("hi"), nil) }
