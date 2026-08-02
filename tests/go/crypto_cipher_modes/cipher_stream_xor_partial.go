// vybe-test: go/crypto_cipher_modes/cipher_stream_xor_partial
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
s := cipher.NewCTR(block, make([]byte, 16))
dst := make([]byte, 5)
src := []byte("hello")
s.XORKeyStream(dst, src) }
