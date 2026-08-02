// vybe-test: go/crypto_cipher_modes/cipher_new_ofb_stream
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
stream := cipher.NewOFB(block, make([]byte, 16))
dst := make([]byte, 32)
stream.XORKeyStream(dst, dst) }
