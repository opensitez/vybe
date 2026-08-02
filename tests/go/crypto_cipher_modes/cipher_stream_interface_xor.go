// vybe-test: go/crypto_cipher_modes/cipher_stream_interface_xor
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
var s cipher.Stream = cipher.NewCTR(block, make([]byte, 16))
buf := make([]byte, 8)
s.XORKeyStream(buf, buf) }
