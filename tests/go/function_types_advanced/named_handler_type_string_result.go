// vybe-test: go/function_types_advanced/named_handler_type_string_result
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type Handler func() string
func run(h Handler) string { return h() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(run(Handler(func() string { return "ok" }))), "ok") }
