// vybe-test: go/switch_fallthrough_extended/fallthrough_with_bool_switch
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch true { case true: __check(fmt.Sprint(1), "1")
fallthrough
case false: __check(fmt.Sprint(0), "0") } }
