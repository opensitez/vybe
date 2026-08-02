// vybe-test: go/constants_iota_advanced/iota_int64_large_shift
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( T0 int64 = 1 << iota; T1; T2 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(T0), "1")
__check(fmt.Sprint(T2), "4") }
