// vybe-test: go/interfaces_patterns_extra/interface_value_roundtrip_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
func wrap(v interface{}) interface{} { return v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(wrap(9)), "9")
}
