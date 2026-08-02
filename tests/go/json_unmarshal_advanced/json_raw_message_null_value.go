// vybe-test: go/json_unmarshal_advanced/json_raw_message_null_value
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type W struct { Raw json.RawMessage }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var w W
json.Unmarshal([]byte(`{"Raw":null}`), &w)
__check(fmt.Sprint(w.Raw == nil), "true") }
