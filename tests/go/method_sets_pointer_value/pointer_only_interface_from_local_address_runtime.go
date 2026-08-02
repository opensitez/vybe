// vybe-test: go/method_sets_pointer_value/pointer_only_interface_from_local_address_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type mutator interface { mutate() }
type data struct { n int }
func (d *data) mutate() { d.n = 99 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { local := data{n: 1}
var m mutator = &local
m.mutate()
__check(fmt.Sprint(local.n), "99") }
