// vybe-test: go/constants_iota_advanced/iota_or_chain_three_flags
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( F1 = 1 << iota; F2; F3 )
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { mask := F1 | F2
__check(fmt.Sprint(mask), "3")
__check(fmt.Sprint(mask | F3), "7") }
