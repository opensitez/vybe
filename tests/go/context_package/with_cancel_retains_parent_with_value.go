// vybe-test: go/context_package/with_cancel_retains_parent_with_value
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

func main() { parent := context.WithValue(context.Background(), "trace", "t1")
child, _ := context.WithCancel(parent)
__check(fmt.Sprint(child.Value("trace").(string)), "t1") }
