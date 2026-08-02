// vybe-test: go/crypto_cipher_modes/cipher_new_gcm_with_tag_size
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
_, _ = cipher.NewGCMWithTagSize(block, 12) }
