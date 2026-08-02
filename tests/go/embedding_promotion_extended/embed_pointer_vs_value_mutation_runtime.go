// vybe-test: go/embedding_promotion_extended/embed_pointer_vs_value_mutation_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type cell struct { n int }
type wrapValue struct { cell }
type wrapPtr struct { *cell }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { v := wrapValue{cell: cell{n: 1}}
p := wrapPtr{cell: &cell{n: 1}}
v.cell.n = 9
p.n = 8
__check(fmt.Sprint(v.n), "9")
__check(fmt.Sprint(p.n), "8") }
