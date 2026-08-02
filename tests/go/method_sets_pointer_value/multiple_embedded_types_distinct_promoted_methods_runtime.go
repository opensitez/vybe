// vybe-test: go/method_sets_pointer_value/multiple_embedded_types_distinct_promoted_methods_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type north struct{}
func (north) dir() string { return "N" }
type east struct{}
func (east) dir() string { return "E" }
type compass struct { north
east }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := compass{}
__check(fmt.Sprint(c.north.dir()), "N")
__check(fmt.Sprint(c.east.dir()), "E") }
