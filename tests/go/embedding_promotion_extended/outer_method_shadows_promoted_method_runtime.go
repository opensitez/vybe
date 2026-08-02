// vybe-test: go/embedding_promotion_extended/outer_method_shadows_promoted_method_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct{}
func (inner) label() string { return "inner" }
type outer struct { inner }
func (outer) label() string { return "outer" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(outer{}.label()), "outer") }
