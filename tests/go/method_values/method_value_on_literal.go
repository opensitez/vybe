// vybe-test: go/method_values/method_value_on_literal
// origin: languages/go/tests/go/test_method_values.rs

package main
import "fmt"
type counter struct { n int }
func (c counter) twice() int { return c.n * 2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { f := counter{n:3}.twice
__check(fmt.Sprint(f()), "6") }
