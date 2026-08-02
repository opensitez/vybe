// vybe-test: go/flag_parsing_extended/flag_commandline_set_does_not_affect_new_flagset
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

func main() { _ = flag.String("shared", "cmd", "")
fs := flag.NewFlagSet("other", flag.ContinueOnError)
local := fs.String("shared", "local", "")
__check(fmt.Sprint(*local), "local") }
