// vybe-test: go/context_cancel_deadline/context_with_cancel_child_only_parent_still_active
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

func main() { parent, _ := context.WithCancel(context.Background())
child, ccancel := context.WithCancel(parent)
ccancel()
__check(fmt.Sprint(parent.Err() == nil), "true") }
