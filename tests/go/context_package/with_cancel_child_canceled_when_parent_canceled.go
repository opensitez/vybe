// vybe-test: go/context_package/with_cancel_child_canceled_when_parent_canceled
// origin: languages/go/tests/go/test_context_package.rs

package main
import "fmt"
import "context"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { parent, parentCancel := context.WithCancel(context.Background())
child, _ := context.WithCancel(parent)
parentCancel()
__check(fmt.Sprint(child.Err() == context.Canceled), "true") }
