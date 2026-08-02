// vybe-test: go/json_unmarshal_advanced/json_marshal_indent_custom_prefix
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type T struct { N int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b, _ := json.MarshalIndent(T{N: 1}, ">", "  ")
s := string(b)
__check(fmt.Sprint(s[0:1]), ">") }
