// vybe-test: go/interface_embedding_methods/left_right_embedded_interface_methods_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type left interface { side() string }
type right interface { edge() string }
type pair interface { left
right }
type both struct{}
func (both) side() string { return "L" }
func (both) edge() string { return "R" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p pair = both{}
__check(fmt.Sprint(p.side()), "L")
__check(fmt.Sprint(p.edge()), "R") }
