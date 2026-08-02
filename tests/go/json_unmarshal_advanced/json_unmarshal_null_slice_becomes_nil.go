// vybe-test: go/json_unmarshal_advanced/json_unmarshal_null_slice_becomes_nil
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

func main() { var s []int
json.Unmarshal([]byte(`null`), &s)
__check(fmt.Sprint(s == nil), "true") }
