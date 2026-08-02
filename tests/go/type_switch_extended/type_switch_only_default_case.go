// vybe-test: go/type_switch_extended/type_switch_only_default_case
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { default: __check(fmt.Sprint("only-default"), "only-default") } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { tag(1.5) }
