// vybe-test: go/json_marshal/marshal_nil_map_null
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

func main() { var m map[string]int
b, _ := json.Marshal(m)
__check(fmt.Sprint(string(b)), "null") }
