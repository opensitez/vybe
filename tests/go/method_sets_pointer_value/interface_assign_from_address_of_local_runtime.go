// vybe-test: go/method_sets_pointer_value/interface_assign_from_address_of_local_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type saver interface { save(int) }
type disk struct { used int }
func (d *disk) save(v int) { d.used += v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { local := disk{}
var s saver = &local
s.save(6)
__check(fmt.Sprint(local.used), "6") }
