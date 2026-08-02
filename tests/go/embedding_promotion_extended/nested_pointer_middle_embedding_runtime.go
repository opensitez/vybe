// vybe-test: go/embedding_promotion_extended/nested_pointer_middle_embedding_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { count int }
type middle struct { *inner }
type outer struct { middle }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{middle: middle{inner: &inner{count: 15}}}
__check(fmt.Sprint(o.count), "15") }
