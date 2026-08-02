// vybe-test: go/stdlib_math_database/big_int_add
// origin: languages/go/tests/go/test_stdlib_math_database.rs

package main
import "fmt"
import "math/big"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := big.NewInt(10)
b := big.NewInt(7)
__check(fmt.Sprint(a.Add(a, b).String()), "17") }
