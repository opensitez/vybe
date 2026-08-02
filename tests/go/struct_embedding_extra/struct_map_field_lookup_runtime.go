// vybe-test: go/struct_embedding_extra/struct_map_field_lookup_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type bag struct { values map[string]int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := bag{values: map[string]int{"x": 9}}
__check(fmt.Sprint(value.values["x"]), "9")
}
