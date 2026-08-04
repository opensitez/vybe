// vybe-test: go/crypto_hash_compile/sha256_sum_empty
// origin: languages/go/tests/go/test_crypto_hash_compile.rs

package main
import "fmt"
import "crypto/sha256"
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

func main() { sum := sha256.Sum256([]byte{})
__p(fmt.Sprint(len(sum))) 
__check("32")
}
