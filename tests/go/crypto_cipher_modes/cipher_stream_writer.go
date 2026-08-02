// vybe-test: go/crypto_cipher_modes/cipher_stream_writer
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
import "bytes"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
stream := cipher.NewCTR(block, make([]byte, 16))
w := &cipher.StreamWriter{S: stream, W: bytes.NewBuffer(nil)}
_, _ = w.Write([]byte("x")) }
