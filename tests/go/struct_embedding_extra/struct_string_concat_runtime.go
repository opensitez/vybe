// vybe-test: go/struct_embedding_extra/struct_string_concat_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type label struct { prefix string
suffix string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := label{prefix: "vy", suffix: "be"}
__check(fmt.Sprint(value.prefix + value.suffix), "vybe")
}
