// vybe-test: go/json_unmarshal_advanced/json_unmarshal_null_pointer_struct_field
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Box struct { N *int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var b Box
json.Unmarshal([]byte(`{"N":null}`), &b)
__check(fmt.Sprint(b.N == nil), "true") }
