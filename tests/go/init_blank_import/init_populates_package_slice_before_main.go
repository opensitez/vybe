// vybe-test: go/init_blank_import/init_populates_package_slice_before_main
// origin: languages/go/tests/go/test_init_blank_import.rs

package main
import "fmt"
var values []int
func init() { values = append(values, 2, 4) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(len(values)), "2")
__check(fmt.Sprint(values[1]), "4") }
