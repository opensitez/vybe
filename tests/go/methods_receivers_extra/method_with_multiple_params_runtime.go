// vybe-test: go/methods_receivers_extra/method_with_multiple_params_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type counter struct { n int }
func (c counter) add(a int, b int) int { return c.n + a + b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := counter{n: 1}
__check(fmt.Sprint(value.add(2, 3)), "6")
}
