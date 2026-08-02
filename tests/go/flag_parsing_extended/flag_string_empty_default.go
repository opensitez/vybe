// vybe-test: go/flag_parsing_extended/flag_string_empty_default
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

func main() { s := flag.String("s", "", "")
__check(fmt.Sprint(len(*s)), "0") }
