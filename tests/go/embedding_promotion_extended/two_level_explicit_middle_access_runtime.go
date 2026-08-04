// vybe-test: go/embedding_promotion_extended/two_level_explicit_middle_access_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type leaf struct { val int }
type branch struct { leaf }
type trunk struct { branch }
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { t := trunk{branch: branch{leaf: leaf{val: 9}}}
__p(fmt.Sprint(t.branch.leaf.val)) 
__check("9")
}
