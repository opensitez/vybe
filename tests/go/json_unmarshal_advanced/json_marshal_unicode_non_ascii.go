// vybe-test: go/json_unmarshal_advanced/json_marshal_unicode_non_ascii
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

func main() { b, _ := json.Marshal("café")
__check(fmt.Sprint(string(b)), "\"café\"") }
