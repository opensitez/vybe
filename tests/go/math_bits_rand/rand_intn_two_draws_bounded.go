// vybe-test: go/math_bits_rand/rand_intn_two_draws_bounded
// origin: languages/go/tests/go/test_math_bits_rand.rs

package main
import "fmt"
import "math/rand"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := rand.Intn(5)
b := rand.Intn(5)
__check(fmt.Sprint(a >= 0 && b >= 0 && a < 5 && b < 5), "true") }
