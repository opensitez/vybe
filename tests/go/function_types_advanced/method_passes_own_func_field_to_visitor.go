// vybe-test: go/function_types_advanced/method_passes_own_func_field_to_visitor
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type node struct { label string
format func(string) string }
func (n node) show(visitor func(string)) { visitor(n.format(n.label)) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := node{label: "go", format: func(s string) string { return s + "!" }}
value.show(func(s string) { __check(fmt.Sprint(s), "go!") }) }
