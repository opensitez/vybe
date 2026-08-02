// vybe-test: go/crypto_cipher_modes/cipher_stream_writer_close
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
import "bytes"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
s := cipher.NewCTR(block, make([]byte, 16))
w := &cipher.StreamWriter{S: s, W: bytes.NewBuffer(nil)}
_ = w.Close() }
