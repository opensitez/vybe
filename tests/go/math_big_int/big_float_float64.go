// vybe-test: go/math_big_int/big_float_float64
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

func main() { f := big.NewFloat(3.14)
v, _ := f.Float64()
__check(fmt.Sprint(v), "3.14") }
