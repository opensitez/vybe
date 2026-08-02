// vybe-test: go/composite_literal_keys/slice_keyed_sparse_high_index
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{10: 1}
__check(fmt.Sprint(len(s)), "11")
__check(fmt.Sprint(s[10]), "1")
__check(fmt.Sprint(s[0]), "0")
}
