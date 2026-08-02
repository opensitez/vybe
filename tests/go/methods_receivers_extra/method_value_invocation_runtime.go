// vybe-test: go/methods_receivers_extra/method_value_invocation_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type counter struct { n int }
func (c counter) total() int { return c.n }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := counter{n: 7}
fn := value.total
__check(fmt.Sprint(fn()), "7")
}
