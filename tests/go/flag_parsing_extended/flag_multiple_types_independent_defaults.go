// vybe-test: go/flag_parsing_extended/flag_multiple_types_independent_defaults
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

func main() { s := flag.String("s", "a", "")
i := flag.Int("i", 1, "")
b := flag.Bool("b", true, "")
__check(fmt.Sprint(*s) + " " + fmt.Sprint(*i) + " " + fmt.Sprint(*b), "a 1 true") }
