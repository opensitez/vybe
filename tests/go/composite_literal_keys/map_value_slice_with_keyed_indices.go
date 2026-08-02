// vybe-test: go/composite_literal_keys/map_value_slice_with_keyed_indices
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string][]int{"data": {0: 5, 2: 7}}
__check(fmt.Sprint(len(m["data"])), "3")
__check(fmt.Sprint(m["data"][0]), "5")
__check(fmt.Sprint(m["data"][2]), "7")
}
