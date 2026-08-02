// vybe-test: go/context_package/with_value_child_inherits_parent_other_key
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

func main() { parent := context.WithValue(context.Background(), "x", 10)
child := context.WithValue(parent, "y", 20)
__check(fmt.Sprint(child.Value("x").(int)), "10")
__check(fmt.Sprint(child.Value("y").(int)), "20") }
