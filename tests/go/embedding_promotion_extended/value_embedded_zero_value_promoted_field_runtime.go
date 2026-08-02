// vybe-test: go/embedding_promotion_extended/value_embedded_zero_value_promoted_field_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { n int }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var o outer
__check(fmt.Sprint(o.n), "0") }
