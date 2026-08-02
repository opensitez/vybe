// vybe-test: go/flag_parsing_extended/flag_int64_default_before_set
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

func main() { size := flag.Int64("size", 1024, "byte size")
__check(fmt.Sprint(*size), "1024") }
