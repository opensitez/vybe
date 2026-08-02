// vybe-test: go/crypto_cipher_modes/aes_key_schedule_256
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
func main() { block, _ := aes.NewCipher(make([]byte, 32))
_ = block }
