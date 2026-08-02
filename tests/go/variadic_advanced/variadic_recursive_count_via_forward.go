// vybe-test: go/variadic_advanced/variadic_recursive_count_via_forward
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func depth(level int, tags ...string) int { if level == 0 { return len(tags) }
return depth(level-1, tags...) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(depth(2, "a", "b", "c")), "3") }
