// vybe-test: go/json_unmarshal_advanced/json_unmarshal_string_tag_float_as_string
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type F struct { Val float64 `json:",string"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var f F
json.Unmarshal([]byte(`{"Val":"3.14"}`), &f)
__check(fmt.Sprint(int(f.Val*100)), "314") }
