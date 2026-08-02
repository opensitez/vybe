// vybe-test: go/struct_embedding_extra/struct_assignment_copies_value_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type counter struct { n int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { left := counter{n: 3}
right := left
right.n = 8
__check(fmt.Sprint(left.n), "3")
__check(fmt.Sprint(right.n), "8")
}
