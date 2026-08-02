// vybe-test: go/math_big_int/big_rat_add
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

func main() { a := big.NewRat(1, 3)
b := big.NewRat(1, 6)
__check(fmt.Sprint(a.Add(a, b).FloatString(2)), "0.50") }
