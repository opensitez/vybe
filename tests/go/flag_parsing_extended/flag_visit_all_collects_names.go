// vybe-test: go/flag_parsing_extended/flag_visit_all_collects_names
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

func main() { _ = flag.Bool("alpha", false, "")
_ = flag.Bool("beta", false, "")
found := 0
flag.VisitAll(func(f *flag.Flag) { if f.Name() == "alpha" || f.Name() == "beta" { found++ } })
__check(fmt.Sprint(found), "2") }
