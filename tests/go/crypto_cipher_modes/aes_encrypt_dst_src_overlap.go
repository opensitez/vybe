// vybe-test: go/crypto_cipher_modes/aes_encrypt_dst_src_overlap
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
buf := make([]byte, 16)
block.Encrypt(buf, buf) }
