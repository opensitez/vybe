// vybe-test: go/crypto_cipher_modes/aes_ecb_encrypt_loop_two_blocks
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
plain := make([]byte, 32)
cipher := make([]byte, 32)
for i := 0; i < len(plain); i += block.BlockSize() { block.Encrypt(cipher[i:i+16], plain[i:i+16]) } }
