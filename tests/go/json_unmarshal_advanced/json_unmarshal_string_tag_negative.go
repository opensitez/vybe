// vybe-test: go/json_unmarshal_advanced/json_unmarshal_string_tag_negative
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type N struct { Val int `json:",string"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var n N
json.Unmarshal([]byte(`{"Val":"-7"}`), &n)
__check(fmt.Sprint(n.Val), "-7") }
