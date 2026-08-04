// vybe-test: go/math_bits_rand/rand_intn_two_draws_bounded
// origin: languages/go/tests/go/test_math_bits_rand.rs

package main
import "fmt"
import "math/rand"
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

func main() { a := rand.Intn(5)
b := rand.Intn(5)
__p(fmt.Sprint(a >= 0 && b >= 0 && a < 5 && b < 5)) 
__check("true")
}
