// vybe-test: go/context_package/without_cancel_shields_child_from_parent_cancel
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
child := context.WithoutCancel(parent)
parentCancel()
__check(fmt.Sprint(child.Err() == nil), "true") }
