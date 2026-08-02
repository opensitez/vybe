// vybe-test: go/interface_embedding_methods/composite_interface_in_slice_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type mover interface { move() int }
type jumper interface { mover }
type hopper struct { steps int }
func (h hopper) move() int { return h.steps }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []jumper{hopper{steps: 3}}
__check(fmt.Sprint(values[0].move()), "3") }
