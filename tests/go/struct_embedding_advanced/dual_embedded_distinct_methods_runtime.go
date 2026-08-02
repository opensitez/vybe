// vybe-test: go/struct_embedding_advanced/dual_embedded_distinct_methods_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type left struct{}
func (left) side() string { return "L" }
type right struct{}
func (right) edge() string { return "R" }
type pair struct { left
right }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := pair{}
__check(fmt.Sprint(value.side()), "L")
__check(fmt.Sprint(value.edge()), "R")
}
