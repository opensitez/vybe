// vybe-test: go/json_marshal/marshal_struct_dash_omits_field
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
type Config struct { Secret string `json:"-"`
OK bool }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := json.Marshal(Config{Secret: "hidden", OK: true})
__check(fmt.Sprint(string(b)), "{\"OK\":true}") }
