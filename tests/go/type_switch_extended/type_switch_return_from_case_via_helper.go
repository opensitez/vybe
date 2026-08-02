// vybe-test: go/type_switch_extended/type_switch_return_from_case_via_helper
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func size(v interface{}) int { switch v.(type) { case string: return len(v.(string))
case int: return v.(int)
default: return 0 } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(size("abc")), "3")
__check(fmt.Sprint(size(10)), "10") }
