// vybe-test: go/struct_embedding_advanced/nested_explicit_middle_field_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

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

func main() { value := outer{middle: middle{inner: inner{count: 11}}}
__check(fmt.Sprint(value.middle.count), "11")
}
