// vybe-test: go/composite_literal_keys/map_value_struct_partial_keyed_fields
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type point struct { x int
y int
label string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string]point{"p": {y: 9, label: "home"}}
__check(fmt.Sprint(m["p"].y), "9")
__check(fmt.Sprint(m["p"].x), "0")
__check(fmt.Sprint(m["p"].label), "home")
}
