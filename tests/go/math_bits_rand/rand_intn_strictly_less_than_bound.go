// vybe-test: go/math_bits_rand/rand_intn_strictly_less_than_bound
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

func main() { n := rand.Intn(10)
__check(fmt.Sprint(n >= 0 && n < 10), "true") }
