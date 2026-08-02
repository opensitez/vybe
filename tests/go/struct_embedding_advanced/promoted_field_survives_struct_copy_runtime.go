// vybe-test: go/struct_embedding_advanced/promoted_field_survives_struct_copy_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct { count int }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { original := outer{inner: inner{count: 2}}
copy := original
copy.count = 5
__check(fmt.Sprint(original.count), "2")
__check(fmt.Sprint(copy.count), "5")
}
