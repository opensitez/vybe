// vybe-test: go/blank_identifier_extended/blank_import_side_effect_with_local_init
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
import _ "strings"
var ready int
func init() { ready = 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ready), "1") }
