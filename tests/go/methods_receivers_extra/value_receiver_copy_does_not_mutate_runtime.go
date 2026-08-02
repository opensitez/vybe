// vybe-test: go/methods_receivers_extra/value_receiver_copy_does_not_mutate_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type counter struct { n int }
func (c counter) bump() { c.n++ }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := counter{n: 3}
value.bump()
__check(fmt.Sprint(value.n), "3")
}
