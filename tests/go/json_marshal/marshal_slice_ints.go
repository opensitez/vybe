// vybe-test: go/json_marshal/marshal_slice_ints
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

func main() { b, _ := json.Marshal([]int{1, 2, 3})
__check(fmt.Sprint(string(b)), "[1,2,3]") }
