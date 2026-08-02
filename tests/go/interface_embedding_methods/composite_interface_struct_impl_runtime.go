// vybe-test: go/interface_embedding_methods/composite_interface_struct_impl_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type opener interface { open() bool }
type closer interface { close() }
type resource interface { opener
closer }
type file struct { ok bool }
func (f file) open() bool { return f.ok }
func (f file) close() {}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var r resource = file{ok: true}
__check(fmt.Sprint(r.open()), "true") }
