// vybe-test: go/json_unmarshal_advanced/json_marshal_string_tag_zero
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

func main() { b, _ := json.Marshal(N{})
__check(fmt.Sprint(string(b)), "{\"Val\":\"0\"}") }
