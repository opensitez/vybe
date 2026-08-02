// vybe-test: go/reflect_value_runtime/reflect_call_no_arg_method
// origin: languages/go/tests/go/test_reflect_value_runtime.rs

package main
import "fmt"
import "reflect"
type Counter struct { n int }
func (c *Counter) Inc() { c.n++ }
func (c Counter) Get() int { return c.n }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := Counter{}
mv := reflect.ValueOf(&c).MethodByName("Get")
out := mv.Call(nil)
__check(fmt.Sprint(out[0].Int()), "0") }
