// vybe-test: go/struct_embedding_extra/anonymous_struct_in_slice_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []struct { name string }{{name: "vybe"}}
__check(fmt.Sprint(values[0].name), "vybe")
}
