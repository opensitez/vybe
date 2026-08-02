// vybe-test: go/context_cancel_deadline/context_with_value_child_shadows_parent_key
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

func main() { parent := context.WithValue(context.Background(), "k", "old")
child := context.WithValue(parent, "k", "new")
__check(fmt.Sprint(child.Value("k").(string)), "new") }
