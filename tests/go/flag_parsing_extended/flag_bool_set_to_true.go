// vybe-test: go/flag_parsing_extended/flag_bool_set_to_true
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

func main() { verbose := flag.Bool("verbose", false, "")
_ = flag.Set("verbose", "true")
__check(fmt.Sprint(*verbose), "true") }
