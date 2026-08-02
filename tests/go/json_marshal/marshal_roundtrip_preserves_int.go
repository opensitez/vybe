// vybe-test: go/json_marshal/marshal_roundtrip_preserves_int
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { orig := 123
b, _ := json.Marshal(orig)
var back int
json.Unmarshal(b, &back)
__check(fmt.Sprint(back), "123") }
