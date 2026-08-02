// vybe-test: go/struct_embedding_advanced/promoted_method_value_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct { n int }
func (i inner) total() int { return i.n }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{inner: inner{n: 6}}
fn := value.total
__check(fmt.Sprint(fn()), "6")
}
