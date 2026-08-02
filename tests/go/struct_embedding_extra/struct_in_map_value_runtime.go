// vybe-test: go/struct_embedding_extra/struct_in_map_value_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type point struct { x int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]point{"a": {x: 11}}
__check(fmt.Sprint(values["a"].x), "11")
}
