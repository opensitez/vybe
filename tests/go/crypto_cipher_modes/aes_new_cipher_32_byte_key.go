// vybe-test: go/crypto_cipher_modes/aes_new_cipher_32_byte_key
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

func main() { _, err := aes.NewCipher(make([]byte, 32))
__check(fmt.Sprint(err == nil), "true") }
