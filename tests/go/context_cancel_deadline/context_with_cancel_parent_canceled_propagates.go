// vybe-test: go/context_cancel_deadline/context_with_cancel_parent_canceled_propagates
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

func main() { parent, pcancel := context.WithCancel(context.Background())
child, _ := context.WithCancel(parent)
pcancel()
__check(fmt.Sprint(child.Err() == context.Canceled), "true") }
