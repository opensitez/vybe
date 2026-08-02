// vybe-test: go/json_unmarshal_advanced/json_unmarshal_null_pointer_then_value
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
json.Unmarshal([]byte(`{"N":8}`), &b)
__check(fmt.Sprint(*b.N), "8") }
