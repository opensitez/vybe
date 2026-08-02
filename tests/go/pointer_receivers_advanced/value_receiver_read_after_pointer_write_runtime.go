// vybe-test: go/pointer_receivers_advanced/value_receiver_read_after_pointer_write_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type note struct { text string }
func (n *note) set(v string) { n.text = v }
func (n note) read() string { return n.text }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := note{text: "a"}
value.set("b")
__check(fmt.Sprint(value.read()), "b")
}
