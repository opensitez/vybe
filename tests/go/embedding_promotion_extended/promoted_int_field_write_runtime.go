// vybe-test: go/embedding_promotion_extended/promoted_int_field_write_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

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

func main() { o := outer{inner: inner{count: 1}}
o.count = 9
__check(fmt.Sprint(o.count), "9") }
