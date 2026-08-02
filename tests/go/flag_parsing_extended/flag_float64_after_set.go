// vybe-test: go/flag_parsing_extended/flag_float64_after_set
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

func main() { ratio := flag.Float64("ratio", 0.0, "")
_ = flag.Set("ratio", "1.25")
__check(fmt.Sprint(*ratio), "1.25") }
