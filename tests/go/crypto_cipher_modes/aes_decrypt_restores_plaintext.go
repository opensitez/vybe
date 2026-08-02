// vybe-test: go/crypto_cipher_modes/aes_decrypt_restores_plaintext
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

func main() { block, _ := aes.NewCipher([]byte{0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15})
plain := []byte{0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1}
cipher := make([]byte, 16)
block.Encrypt(cipher, plain)
out := make([]byte, 16)
block.Decrypt(out, cipher)
__check(fmt.Sprint(out[15]), "1") }
