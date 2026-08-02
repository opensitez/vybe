// vybe-test: go/embedding_promotion_extended/two_embedded_same_field_name_requires_qualifier_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type left struct { id int }
type right struct { id int }
type pair struct { left
right }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := pair{left: left{id: 1}, right: right{id: 2}}
__check(fmt.Sprint(p.left.id), "1")
__check(fmt.Sprint(p.right.id), "2") }
