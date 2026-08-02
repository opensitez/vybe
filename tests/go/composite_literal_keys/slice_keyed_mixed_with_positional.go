// vybe-test: go/composite_literal_keys/slice_keyed_mixed_with_positional
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1: 10, 20, 30}
__check(fmt.Sprint(s[1]), "10")
__check(fmt.Sprint(s[2]), "20")
__check(fmt.Sprint(s[3]), "30")
__check(fmt.Sprint(len(s)), "4")
}
