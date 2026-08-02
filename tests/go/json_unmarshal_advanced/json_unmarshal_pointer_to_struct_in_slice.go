// vybe-test: go/json_unmarshal_advanced/json_unmarshal_pointer_to_struct_in_slice
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Item struct { N int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s []*Item
json.Unmarshal([]byte(`[{"N":1},{"N":2}]`), &s)
__check(fmt.Sprint(s[1].N), "2") }
