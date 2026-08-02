// vybe-test: go/composite_literal_keys/array_keyed_index_zero_explicit
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := [5]int{0: 11, 4: 44}
__check(fmt.Sprint(a[0]), "11")
__check(fmt.Sprint(a[1]), "0")
__check(fmt.Sprint(a[4]), "44")
}
