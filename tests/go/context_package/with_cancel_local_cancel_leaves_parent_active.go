// vybe-test: go/context_package/with_cancel_local_cancel_leaves_parent_active
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

func main() { parent, _ := context.WithCancel(context.Background())
child, childCancel := context.WithCancel(parent)
childCancel()
__check(fmt.Sprint(parent.Err() == nil), "true")
__check(fmt.Sprint(child.Err() == context.Canceled), "true") }
