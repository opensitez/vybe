// vybe-test: go/flag_parsing_extended/flag_new_flagset_visit_all_after_define
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

func main() { fs := flag.NewFlagSet("fs", flag.ContinueOnError)
_ = fs.Bool("on", false, "")
n := 0
fs.VisitAll(func(f *flag.Flag) { n++ })
__check(fmt.Sprint(n), "1") }
