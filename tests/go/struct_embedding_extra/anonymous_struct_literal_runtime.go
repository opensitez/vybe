// vybe-test: go/struct_embedding_extra/anonymous_struct_literal_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := struct { left int
right int }{left: 2, right: 8}
__check(fmt.Sprint(value.left + value.right), "10")
}
