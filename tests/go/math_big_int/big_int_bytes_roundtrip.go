// vybe-test: go/math_big_int/big_int_bytes_roundtrip
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

func main() { orig := big.NewInt(1000)
back := new(big.Int).SetBytes(orig.Bytes())
__check(fmt.Sprint(back.String()), "1000") }
