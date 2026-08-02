// vybe-test: go/crypto_cipher_modes/cipher_stream_reader
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
import "io"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
stream := cipher.NewCTR(block, make([]byte, 16))
r := &cipher.StreamReader{S: stream, R: nil}
_ = r }
