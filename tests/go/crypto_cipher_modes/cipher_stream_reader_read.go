// vybe-test: go/crypto_cipher_modes/cipher_stream_reader_read
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
import "bytes"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
s := cipher.NewCTR(block, make([]byte, 16))
r := &cipher.StreamReader{S: s, R: bytes.NewReader([]byte("data"))}
buf := make([]byte, 4)
_, _ = r.Read(buf) }
