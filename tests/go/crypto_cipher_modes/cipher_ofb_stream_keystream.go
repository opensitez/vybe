// vybe-test: go/crypto_cipher_modes/cipher_ofb_stream_keystream
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
s := cipher.NewOFB(block, make([]byte, 16))
buf := make([]byte, 64)
s.XORKeyStream(buf, buf) }
