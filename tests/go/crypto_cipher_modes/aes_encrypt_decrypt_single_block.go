// vybe-test: go/crypto_cipher_modes/aes_encrypt_decrypt_single_block
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs

package main
import "fmt"
import "crypto/aes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { key := make([]byte, 16)
block, _ := aes.NewCipher(key)
plain := make([]byte, 16)
cipher := make([]byte, 16)
block.Encrypt(cipher, plain)
out := make([]byte, 16)
block.Decrypt(out, cipher)
__check(fmt.Sprint(out[0] == plain[0]), "true") }
