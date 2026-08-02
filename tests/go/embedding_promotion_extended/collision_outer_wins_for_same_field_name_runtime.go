// vybe-test: go/embedding_promotion_extended/collision_outer_wins_for_same_field_name_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type base struct { score int }
type derived struct { base
score int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { d := derived{base: base{score: 1}, score: 2}
__check(fmt.Sprint(d.score), "2")
__check(fmt.Sprint(d.base.score), "1") }
