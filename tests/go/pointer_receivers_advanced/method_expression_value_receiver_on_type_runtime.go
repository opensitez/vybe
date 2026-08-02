// vybe-test: go/pointer_receivers_advanced/method_expression_value_receiver_on_type_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type tag struct { name string }
func (t tag) label() string { return t.name }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fn := tag.label
__check(fmt.Sprint(fn(tag{name: "go"})), "go")
}
