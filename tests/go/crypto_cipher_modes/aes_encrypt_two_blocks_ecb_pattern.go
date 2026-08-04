// vybe-test: go/crypto_cipher_modes/aes_encrypt_two_blocks_ecb_pattern
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

func main() { block, _ := aes.NewCipher(make([]byte, 16))
p1 := make([]byte, 16)
p2 := make([]byte, 16)
p2[0] = 1
c1 := make([]byte, 16)
c2 := make([]byte, 16)
block.Encrypt(c1, p1)
block.Encrypt(c2, p2)
__p(fmt.Sprint(c1[0] != c2[0])) 
__check("true")
}
