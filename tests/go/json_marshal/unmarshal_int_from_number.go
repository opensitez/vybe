// vybe-test: go/json_marshal/unmarshal_int_from_number
// origin: languages/go/tests/go/test_json_marshal.rs

package main
import "fmt"
import "encoding/json"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var n int
json.Unmarshal([]byte("42"), &n)
__check(fmt.Sprint(n), "42") }
