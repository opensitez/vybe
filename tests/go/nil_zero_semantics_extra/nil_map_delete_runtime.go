// vybe-test: go/nil_zero_semantics_extra/nil_map_delete_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var values map[string]int
delete(values, "x")
__check(fmt.Sprint(values == nil), "true")
}
