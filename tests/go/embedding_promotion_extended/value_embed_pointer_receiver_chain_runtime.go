// vybe-test: go/embedding_promotion_extended/value_embed_pointer_receiver_chain_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type engine struct { rpm int }
func (e *engine) rev() *engine { e.rpm++
return e }
type car struct { engine }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := car{engine: engine{rpm: 100}}
c.rev().rev()
__check(fmt.Sprint(c.rpm), "102") }
