// vybe-test: go/json_marshal/marshal_nested_object
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
type Child struct { N int }
type Parent struct { Child Child
Tag string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := json.Marshal(Parent{Child: Child{N: 1}, Tag: "x"})
__check(fmt.Sprint(string(b)), "{\"Child\":{\"N\":1},\"Tag\":\"x\"}") }
