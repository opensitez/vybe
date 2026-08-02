// vybe-test: go/init_blank_import/init_calls_package_helper_before_main
// origin: languages/go/tests/go/test_init_blank_import.rs

package main
import "fmt"
var ready bool
func markReady() { ready = true }
func init() { markReady() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ready), "true") }
