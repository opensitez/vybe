// vybe-test: go/json_unmarshal_advanced/json_unmarshal_null_pointer_slice_element
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

func main() { var s []*int
json.Unmarshal([]byte(`[1,null,3]`), &s)
__check(fmt.Sprint(s[0] != nil), "true")
__check(fmt.Sprint(s[1] == nil), "true")
__check(fmt.Sprint(*s[2]), "3") }
