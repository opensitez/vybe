// vybe-test: go/json_marshal/marshal_struct_renames_with_tag
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

func main() { b, _ := json.Marshal(Item{ID: 1})
__check(fmt.Sprint(string(b)), "{\"id\":1}") }
