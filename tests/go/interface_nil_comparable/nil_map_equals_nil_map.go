// vybe-test: go/interface_nil_comparable/nil_map_equals_nil_map
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var left map[string]int
var right map[string]int
__check(fmt.Sprint(left == nil), "true")
__check(fmt.Sprint(left == right), "true") }
