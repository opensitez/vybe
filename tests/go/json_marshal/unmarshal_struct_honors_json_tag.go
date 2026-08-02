// vybe-test: go/json_marshal/unmarshal_struct_honors_json_tag
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
type Item struct { ID int `json:"id"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var it Item
json.Unmarshal([]byte("{\"id\":99}"), &it)
__check(fmt.Sprint(it.ID), "99") }
