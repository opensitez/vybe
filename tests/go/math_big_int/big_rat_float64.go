// vybe-test: go/math_big_int/big_rat_float64
// origin: languages/go/tests/go/test_math_big_int.rs

package main
import "fmt"
import "math/big"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := big.NewRat(1, 4)
f, _ := r.Float64()
__check(fmt.Sprint(f), "0.25") }
