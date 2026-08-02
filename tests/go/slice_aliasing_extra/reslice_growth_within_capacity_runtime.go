// vybe-test: go/slice_aliasing_extra/reslice_growth_within_capacity_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []int{1, 2, 3, 4}
part := values[:2]
part = part[:4]
__check(fmt.Sprint(part[3]), "4")
}
