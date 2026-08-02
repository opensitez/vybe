// vybe-test: go/json_unmarshal_advanced/json_unmarshal_unicode_escape_in_array
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

func main() { var s []string
json.Unmarshal([]byte(`["\u0061","\u0062"]`), &s)
__check(fmt.Sprint(s[0]), "a")
__check(fmt.Sprint(s[1]), "b") }
