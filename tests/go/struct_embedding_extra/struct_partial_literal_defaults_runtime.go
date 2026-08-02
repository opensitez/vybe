// vybe-test: go/struct_embedding_extra/struct_partial_literal_defaults_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type item struct { count int
name string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := item{name: "go"}
__check(fmt.Sprint(value.count), "0")
__check(fmt.Sprint(value.name), "go")
}
