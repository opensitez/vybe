// vybe-test: go/function_types_advanced/struct_two_func_fields_distinct_results
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type pair struct { left func() int
right func() int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := pair{left: func() int { return 1 }, right: func() int { return 2 }}
__check(fmt.Sprint(value.left()), "1")
__check(fmt.Sprint(value.right()), "2") }
