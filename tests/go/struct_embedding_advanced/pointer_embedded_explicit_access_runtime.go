// vybe-test: go/struct_embedding_advanced/pointer_embedded_explicit_access_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct { count int }
type outer struct { *inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{inner: &inner{count: 8}}
__check(fmt.Sprint(value.inner.count), "8")
}
