// vybe-test: go/json_unmarshal_advanced/json_unmarshal_unicode_escape_cyrillic
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

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
json.Unmarshal([]byte(`"\u043f\u0440\u0438\u0432\u0435\u0442"`), &s)
__check(fmt.Sprint(len(s)), "12") }
