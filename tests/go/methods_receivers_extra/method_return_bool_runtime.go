// vybe-test: go/methods_receivers_extra/method_return_bool_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type gate struct { open bool }
func (g gate) ready() bool { return g.open }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(gate{open: false}.ready()), "false")
}
