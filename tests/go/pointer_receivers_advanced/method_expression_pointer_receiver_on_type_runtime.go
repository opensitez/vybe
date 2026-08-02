// vybe-test: go/pointer_receivers_advanced/method_expression_pointer_receiver_on_type_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type tag struct { name string }
func (t *tag) rename(v string) { t.name = v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := &tag{name: "old"}
fn := (*tag).rename
fn(value, "new")
__check(fmt.Sprint(value.name), "new")
}
