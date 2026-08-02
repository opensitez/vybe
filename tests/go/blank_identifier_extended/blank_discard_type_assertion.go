// vybe-test: go/blank_identifier_extended/blank_discard_type_assertion
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = 42
_, ok := v.(int)
__check(fmt.Sprint(ok), "true") }
