// vybe-test: go/context_package/with_value_child_overrides_parent_same_key
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

func main() { parent := context.WithValue(context.Background(), "k", 1)
child := context.WithValue(parent, "k", 2)
__check(fmt.Sprint(child.Value("k").(int)), "2") }
