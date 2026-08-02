// vybe-test: go/json_marshal/marshal_map_one_entry
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := json.Marshal(map[string]int{"a": 1})
__check(fmt.Sprint(string(b)), "{\"a\":1}") }
