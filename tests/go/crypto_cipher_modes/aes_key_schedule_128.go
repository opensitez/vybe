// vybe-test: go/crypto_cipher_modes/aes_key_schedule_128
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs
// vybe-test-mode: compile

package main
import "crypto/aes"
func main() { block, _ := aes.NewCipher([]byte{0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15})
_ = block }
