// vybe-test: go/composite_literal_keys/nested_map_array_struct_keyed_chain
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type item struct { id int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { data := map[string][]item{"batch": {{id: 10}, {id: 20}}}
__check(fmt.Sprint(data["batch"][0].id), "10")
__check(fmt.Sprint(data["batch"][1].id), "20")
}
