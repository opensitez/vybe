// vybe-test: go/maphash_crypto_rand/maphash_string_sum
// origin: languages/go/tests/go/test_maphash_crypto_rand.rs

package main
import "fmt"
import "hash/maphash"
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

func main() { var h maphash.Hash
h.SetSeed(maphash.MakeSeed())
h.WriteString("vybe")
__p(fmt.Sprint(h.Sum64() > 0)) 
__check("true")
}
