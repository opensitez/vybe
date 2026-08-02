// vybe-test: go/struct_embedding_extra/struct_copy_before_mutation_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type item struct { n int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { original := item{n: 4}
copy := original
original.n = 10
__check(fmt.Sprint(copy.n), "4")
__check(fmt.Sprint(original.n), "10")
}
