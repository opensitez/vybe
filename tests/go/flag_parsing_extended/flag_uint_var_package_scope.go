// vybe-test: go/flag_parsing_extended/flag_uint_var_package_scope
// origin: languages/go/tests/go/test_flag_parsing_extended.rs

package main
import "fmt"
import "flag"
var shards = flag.Uint("shards", 3, "")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(*shards), "3") }
