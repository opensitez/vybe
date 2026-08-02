// vybe-test: go/crypto_cipher_modes/aes_encrypt_two_blocks_ecb_pattern
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

func main() { block, _ := aes.NewCipher(make([]byte, 16))
p1 := make([]byte, 16)
p2 := make([]byte, 16)
p2[0] = 1
c1 := make([]byte, 16)
c2 := make([]byte, 16)
block.Encrypt(c1, p1)
block.Encrypt(c2, p2)
__check(fmt.Sprint(c1[0] != c2[0]), "true") }
