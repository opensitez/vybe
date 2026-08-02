// vybe-test: go/composite_literal_keys/slice_of_maps_with_keyed_struct_values
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type pair struct { a int
b int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []map[string]pair{{"x": {b: 2, a: 1}}, {"y": pair{a: 3, b: 4}}}
__check(fmt.Sprint(s[0]["x"].a), "1")
__check(fmt.Sprint(s[1]["y"].b), "4")
}
