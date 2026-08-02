// vybe-test: go/function_types_advanced/struct_func_field_reassigned_and_called
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type holder struct { fn func(string) string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := holder{}
__check(fmt.Sprint(value.fn == nil), "true")
value.fn = func(s string) string { return s + "!" }
__check(fmt.Sprint(value.fn("go")), "go!") }
