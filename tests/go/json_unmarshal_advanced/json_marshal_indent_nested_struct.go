// vybe-test: go/json_unmarshal_advanced/json_marshal_indent_nested_struct
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Inner struct { V int }
type Outer struct { Inner Inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := json.MarshalIndent(Outer{Inner: Inner{V: 2}}, "", "  ")
__check(fmt.Sprint(len(b) > 10), "true") }
