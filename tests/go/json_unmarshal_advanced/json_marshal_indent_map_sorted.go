// vybe-test: go/json_unmarshal_advanced/json_marshal_indent_map_sorted
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := json.MarshalIndent(map[string]int{"b": 2, "a": 1}, "", "  ")
__check(fmt.Sprint(len(b) > 5), "true") }
