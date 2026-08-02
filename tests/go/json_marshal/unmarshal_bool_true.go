// vybe-test: go/json_marshal/unmarshal_bool_true
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

func main() { var v bool
json.Unmarshal([]byte("true"), &v)
__check(fmt.Sprint(v), "true") }
