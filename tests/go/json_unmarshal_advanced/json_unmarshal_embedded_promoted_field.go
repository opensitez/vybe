// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_promoted_field
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Inner struct { N int }
type Outer struct { Inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var o Outer
json.Unmarshal([]byte(`{"N":5}`), &o)
__check(fmt.Sprint(o.N), "5") }
