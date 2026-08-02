// vybe-test: go/json_marshal/unmarshal_slice_of_ints
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

func main() { var s []int
json.Unmarshal([]byte("[10,20,30]"), &s)
__check(fmt.Sprint(len(s)), "3")
__check(fmt.Sprint(s[1]), "20") }
