// vybe-test: go/slice_aliasing_extra/subslice_observes_source_mutation_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []int{1, 2, 3}
part := values[:2]
values[0] = 7
__check(fmt.Sprint(part[0]), "7")
}
