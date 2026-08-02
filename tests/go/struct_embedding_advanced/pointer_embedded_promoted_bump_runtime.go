// vybe-test: go/struct_embedding_advanced/pointer_embedded_promoted_bump_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct { n int }
func (i *inner) bump() { i.n++ }
type outer struct { *inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{inner: &inner{n: 2}}
value.bump()
__check(fmt.Sprint(value.n), "3")
}
