// vybe-test: go/crypto_cipher_modes/cipher_block_interface
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
import "crypto/cipher"
func main() { var b cipher.Block
b, _ = aes.NewCipher(make([]byte, 16))
_ = b.BlockSize() }
