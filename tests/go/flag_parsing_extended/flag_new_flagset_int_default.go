// vybe-test: go/flag_parsing_extended/flag_new_flagset_int_default
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

func main() { fs := flag.NewFlagSet("tool", flag.ContinueOnError)
n := fs.Int("n", 7, "")
__check(fmt.Sprint(*n), "7") }
