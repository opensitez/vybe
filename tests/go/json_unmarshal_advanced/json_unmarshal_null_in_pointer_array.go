// vybe-test: go/json_unmarshal_advanced/json_unmarshal_null_in_pointer_array
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

func main() { var s []*string
json.Unmarshal([]byte(`["a",null,"c"]`), &s)
__check(fmt.Sprint(*s[0]), "a")
__check(fmt.Sprint(s[1] == nil), "true")
__check(fmt.Sprint(*s[2]), "c") }
