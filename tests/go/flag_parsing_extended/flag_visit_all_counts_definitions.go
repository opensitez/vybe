// vybe-test: go/flag_parsing_extended/flag_visit_all_counts_definitions
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

func main() { _ = flag.String("a", "", "")
_ = flag.Int("b", 0, "")
n := 0
flag.VisitAll(func(f *flag.Flag) { n++ })
__check(fmt.Sprint(n >= 2), "true") }
