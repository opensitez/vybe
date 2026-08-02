// vybe-test: go/struct_embedding_extra/empty_struct_slice_len_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type token struct{}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []token{{}, {}}
__check(fmt.Sprint(len(values)), "2")
}
