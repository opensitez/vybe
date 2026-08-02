// vybe-test: go/struct_embedding_extra/struct_function_field_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type holder struct { fn func(int) int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := holder{fn: func(v int) int { return v + 3 }}
__check(fmt.Sprint(value.fn(4)), "7")
}
