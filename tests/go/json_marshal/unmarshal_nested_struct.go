// vybe-test: go/json_marshal/unmarshal_nested_struct
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
type Child struct { N int }
type Parent struct { Child Child }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p Parent
json.Unmarshal([]byte("{\"Child\":{\"N\":5}}"), &p)
__check(fmt.Sprint(p.Child.N), "5") }
