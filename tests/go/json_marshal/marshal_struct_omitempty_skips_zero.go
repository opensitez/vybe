// vybe-test: go/json_marshal/marshal_struct_omitempty_skips_zero
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
type Data struct { Count int `json:",omitempty"`
Label string `json:",omitempty"` }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := json.Marshal(Data{})
__check(fmt.Sprint(string(b)), "{}") }
