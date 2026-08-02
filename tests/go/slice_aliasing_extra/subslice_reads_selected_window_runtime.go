// vybe-test: go/slice_aliasing_extra/subslice_reads_selected_window_runtime
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
part := values[1:]
__check(fmt.Sprint(part[0]), "2")
__check(fmt.Sprint(len(part)), "2")
}
