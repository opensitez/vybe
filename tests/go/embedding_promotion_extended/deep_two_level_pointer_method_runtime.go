// vybe-test: go/embedding_promotion_extended/deep_two_level_pointer_method_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type leaf struct { n int }
func (l *leaf) inc() { l.n++ }
type branch struct { leaf }
type trunk struct { branch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := trunk{branch: branch{leaf: leaf{n: 0}}}
t.inc()
__check(fmt.Sprint(t.n), "1") }
