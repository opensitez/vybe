// vybe-test: go/json_unmarshal_advanced/json_unmarshal_empty_object_to_struct
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type S struct { X int
Y string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s S
json.Unmarshal([]byte(`{}`), &s)
__check(fmt.Sprint(s.X), "0")
__check(fmt.Sprint(s.Y == ""), "true") }
