// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_pointer_nil
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Inner struct { N int }
type Outer struct { *Inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var o Outer
json.Unmarshal([]byte(`{"N":3}`), &o)
__check(fmt.Sprint(o.Inner.N), "3") }
