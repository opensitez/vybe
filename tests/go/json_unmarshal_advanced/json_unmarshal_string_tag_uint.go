// vybe-test: go/json_unmarshal_advanced/json_unmarshal_string_tag_uint
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type U struct { Val uint `json:",string"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var u U
json.Unmarshal([]byte(`{"Val":"255"}`), &u)
__check(fmt.Sprint(u.Val), "255") }
