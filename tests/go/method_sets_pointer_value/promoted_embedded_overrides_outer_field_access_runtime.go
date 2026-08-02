// vybe-test: go/method_sets_pointer_value/promoted_embedded_overrides_outer_field_access_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type inner struct { x int }
type outer struct { inner
x int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{x: 1}, x: 2}
__check(fmt.Sprint(o.x), "2")
__check(fmt.Sprint(o.inner.x), "1") }
