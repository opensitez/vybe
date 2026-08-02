// vybe-test: go/composite_literal_keys/array_keyed_mixed_with_positional_continuation
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := [6]int{1: 10, 3: 30, 5}
__check(fmt.Sprint(a[1]), "10")
__check(fmt.Sprint(a[3]), "30")
__check(fmt.Sprint(a[4]), "5")
__check(fmt.Sprint(len(a)), "6")
}
