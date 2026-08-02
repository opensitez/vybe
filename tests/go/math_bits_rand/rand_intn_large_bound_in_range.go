// vybe-test: go/math_bits_rand/rand_intn_large_bound_in_range
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

func main() { __check(fmt.Sprint(rand.Intn(1000000) < 1000000), "true") }
