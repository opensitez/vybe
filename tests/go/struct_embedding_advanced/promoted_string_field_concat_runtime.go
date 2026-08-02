// vybe-test: go/struct_embedding_advanced/promoted_string_field_concat_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct { prefix string
suffix string }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{inner: inner{prefix: "go", suffix: "lang"}}
__check(fmt.Sprint(value.prefix + value.suffix), "golang")
}
