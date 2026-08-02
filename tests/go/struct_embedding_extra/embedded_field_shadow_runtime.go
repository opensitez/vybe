// vybe-test: go/struct_embedding_extra/embedded_field_shadow_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type inner struct { name string }
type outer struct { inner
name string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{inner: inner{name: "inner"}, name: "outer"}
__check(fmt.Sprint(value.name), "outer")
__check(fmt.Sprint(value.inner.name), "inner")
}
