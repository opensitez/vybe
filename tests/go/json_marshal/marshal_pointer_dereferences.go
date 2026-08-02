// vybe-test: go/json_marshal/marshal_pointer_dereferences
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

func main() { n := 9
b, _ := json.Marshal(&n)
__check(fmt.Sprint(string(b)), "9") }
