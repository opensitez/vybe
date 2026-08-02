// vybe-test: go/composite_literal_keys/map_nested_struct_and_slice_keys
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type cell struct { n int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string]struct { rows []cell }{ "t": {rows: []cell{{n: 1}, {n: 2}}} }
__check(fmt.Sprint(m["t"].rows[1].n), "2")
}
