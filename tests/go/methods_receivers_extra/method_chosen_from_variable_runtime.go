// vybe-test: go/methods_receivers_extra/method_chosen_from_variable_runtime
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

func main() { value := counter{n: 16}
other := value
__check(fmt.Sprint(other.total()), "16")
}
