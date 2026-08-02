// vybe-test: go/context_cancel_deadline/context_with_value_parent_visible_from_child
// origin: languages/go/tests/go/test_context_cancel_deadline.rs

package main
import "fmt"
import "context"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { parent := context.WithValue(context.Background(), "trace", "root")
child, _ := context.WithCancel(parent)
__check(fmt.Sprint(child.Value("trace").(string)), "root") }
