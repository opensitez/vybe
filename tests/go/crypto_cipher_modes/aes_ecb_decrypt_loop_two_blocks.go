// vybe-test: go/crypto_cipher_modes/aes_ecb_decrypt_loop_two_blocks
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
func main() { block, _ := aes.NewCipher(make([]byte, 16))
cipher := make([]byte, 32)
plain := make([]byte, 32)
for i := 0; i < len(cipher); i += block.BlockSize() { block.Decrypt(plain[i:i+16], cipher[i:i+16]) } }
