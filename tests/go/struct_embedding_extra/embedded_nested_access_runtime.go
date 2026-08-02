// vybe-test: go/struct_embedding_extra/embedded_nested_access_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type inner struct { count int }
type middle struct { inner }
type outer struct { middle }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{middle: middle{inner: inner{count: 12}}}
__check(fmt.Sprint(value.count), "12")
}
