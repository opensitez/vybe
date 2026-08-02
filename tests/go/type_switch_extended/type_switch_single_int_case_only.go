// vybe-test: go/type_switch_extended/type_switch_single_int_case_only
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int: __check(fmt.Sprint("int"), "int") } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { tag(7) }
