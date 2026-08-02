// vybe-test: go/flag_parsing_extended/flag_bool_set_literal_0
// origin: languages/go/tests/go/test_flag_parsing_extended.rs

package main
import "fmt"
import "flag"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b := flag.Bool("b", true, "")
_ = flag.Set("b", "0")
__check(fmt.Sprint(*b), "false") }
