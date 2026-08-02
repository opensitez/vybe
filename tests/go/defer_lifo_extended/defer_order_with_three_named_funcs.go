// vybe-test: go/defer_lifo_extended/defer_order_with_three_named_funcs
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func p1() { __check(fmt.Sprint(1), "3") }
func p2() { __check(fmt.Sprint(2), "2") }
func p3() { __check(fmt.Sprint(3), "1") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer p1()
defer p2()
defer p3()
}
