// vybe-test: go/function_types_advanced/struct_passes_field_func_to_helper
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type box struct { transform func(int) int }
func invoke(b box, v int) int { return b.transform(v) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := box{transform: func(v int) int { return v - 1 }}
__check(fmt.Sprint(invoke(value, 9)), "8") }
