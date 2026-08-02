// vybe-test: go/json_unmarshal_advanced/json_unmarshal_string_tag_bool
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type B struct { Ok bool `json:",string"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var b B
json.Unmarshal([]byte(`{"Ok":"true"}`), &b)
__check(fmt.Sprint(b.Ok), "true") }
