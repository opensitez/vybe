// vybe-test: go/embedding_promotion_extended/two_level_explicit_middle_access_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type leaf struct { val int }
type branch struct { leaf }
type trunk struct { branch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := trunk{branch: branch{leaf: leaf{val: 9}}}
__check(fmt.Sprint(t.branch.leaf.val), "9") }
