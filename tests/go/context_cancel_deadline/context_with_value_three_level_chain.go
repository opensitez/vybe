// vybe-test: go/context_cancel_deadline/context_with_value_three_level_chain
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

func main() { c1 := context.WithValue(context.Background(), "a", 1)
c2 := context.WithValue(c1, "b", 2)
c3 := context.WithValue(c2, "c", 3)
__check(fmt.Sprint(c3.Value("a").(int)), "1")
__check(fmt.Sprint(c3.Value("b").(int)), "2")
__check(fmt.Sprint(c3.Value("c").(int)), "3") }
