// vybe-test: go/interface_embedding_methods/pointer_receiver_embedded_interface_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type scaler interface { scale() int }
type resizable interface { scaler }
type icon struct { n int }
func (i *icon) scale() int { return i.n * 2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var r resizable = &icon{n: 5}
__check(fmt.Sprint(r.scale()), "10") }
