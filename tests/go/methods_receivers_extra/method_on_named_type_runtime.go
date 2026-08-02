// vybe-test: go/methods_receivers_extra/method_on_named_type_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type score int
func (s score) next() int { return int(s) + 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value score = 8
__check(fmt.Sprint(value.next()), "9")
}
