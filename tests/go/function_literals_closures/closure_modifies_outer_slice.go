// vybe-test: go/function_literals_closures/closure_modifies_outer_slice
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { items := []int{}
push := func(v int) { items = append(items, v) }
push(1)
push(2)
__check(fmt.Sprint(len(items)), "2")
__check(fmt.Sprint(items[1]), "2") }
