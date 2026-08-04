// vybe-test: go/crypto_cipher_modes/aes_encrypt_decrypt_single_block
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs

package main
import "fmt"
import "crypto/aes"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
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
__p(fmt.Sprint(out[0] == plain[0])) 
__check("true")
}
