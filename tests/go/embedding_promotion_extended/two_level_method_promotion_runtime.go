// vybe-test: go/embedding_promotion_extended/two_level_method_promotion_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type leaf struct{}
func (leaf) deep() string { return "L" }
type branch struct { leaf }
type trunk struct { branch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(trunk{}.deep()), "L") }
