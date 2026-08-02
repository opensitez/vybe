// vybe-test: go/json_marshal/marshal_nil_pointer_null
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

func main() { var p *int
b, _ := json.Marshal(p)
__check(fmt.Sprint(string(b)), "null") }
