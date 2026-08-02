// vybe-test: go/json_marshal/unmarshal_string_from_quoted
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

func main() { var s string
json.Unmarshal([]byte("\"hi\""), &s)
__check(fmt.Sprint(s), "hi") }
