// vybe-test: go/embedding_promotion_extended/embedded_string_field_concat_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { left string
right string }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{left: "go", right: "lang"}}
__check(fmt.Sprint(o.left + o.right), "golang") }
