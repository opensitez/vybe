// vybe-test: go/crypto_cipher_modes/cipher_new_cfb_decrypter_stream
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
stream := cipher.NewCFBDecrypter(block, make([]byte, 16))
dst := make([]byte, 16)
stream.XORKeyStream(dst, dst) }
