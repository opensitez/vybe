// vybe-test: go/methods_receivers_extra/method_on_embedded_field_explicit_runtime
// origin: languages/go/tests/go/test_methods_receivers_extra.rs

package main
import "fmt"
type inner struct{}
func (inner) label() string { return "ok" }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{}
__check(fmt.Sprint(value.inner.label()), "ok")
}
