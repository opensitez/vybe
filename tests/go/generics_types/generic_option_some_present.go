// vybe-test: go/generics_types/generic_option_some_present
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Option[T any] struct { Value T
Present bool }
func Some[T any](v T) Option[T] { return Option[T]{Value: v, Present: true} }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := Some(42)
__check(fmt.Sprint(o.Present), "true")
__check(fmt.Sprint(o.Value), "42") }
