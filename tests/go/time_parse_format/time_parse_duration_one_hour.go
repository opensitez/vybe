// vybe-test: go/time_parse_format/time_parse_duration_one_hour
// origin: languages/go/tests/go/test_time_parse_format.rs

package main
import "fmt"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { d, _ := time.ParseDuration("1h")
__check(fmt.Sprint(d.Hours()), "1") }
