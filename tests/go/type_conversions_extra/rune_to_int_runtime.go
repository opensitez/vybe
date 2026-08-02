// vybe-test: go/type_conversions_extra/rune_to_int_runtime
// origin: languages/go/tests/go/test_type_conversions_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value rune = 'A'
__check(fmt.Sprint(int(value)), "65")
}
