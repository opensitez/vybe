// vybe-test: go/interface_embedding_methods/promoted_method_with_int_arg_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type counter interface { bump(int) int }
type meter interface { counter }
type gauge struct { n int }
func (g gauge) bump(delta int) int { return g.n + delta }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m meter = gauge{n: 4}
__check(fmt.Sprint(m.bump(3)), "7") }
