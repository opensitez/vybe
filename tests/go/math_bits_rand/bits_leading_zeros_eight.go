// vybe-test: go/math_bits_rand/bits_leading_zeros_eight
// origin: languages/go/tests/go/test_math_bits_rand.rs

package main
import "fmt"
import "math/bits"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(bits.LeadingZeros(8)), "60") }
