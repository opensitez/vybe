// vybe-test: go/flag_parsing_extended/flag_lookup_name_matches
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

func main() { _ = flag.Int("port", 80, "")
f := flag.Lookup("port")
__check(fmt.Sprint(f.Name()), "port") }
