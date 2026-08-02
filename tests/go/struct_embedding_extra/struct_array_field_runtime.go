// vybe-test: go/struct_embedding_extra/struct_array_field_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type bag struct { values [3]int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := bag{values: [3]int{1, 2, 3}}
__check(fmt.Sprint(value.values[2]), "3")
}
