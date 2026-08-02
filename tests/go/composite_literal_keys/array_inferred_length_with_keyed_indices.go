// vybe-test: go/composite_literal_keys/array_inferred_length_with_keyed_indices
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := [...]int{3: 9, 4: 10}
__check(fmt.Sprint(len(a)), "5")
__check(fmt.Sprint(a[3]), "9")
__check(fmt.Sprint(a[0]), "0")
}
