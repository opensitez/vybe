// vybe-test: go/time_parse_format/time_duration_nanoseconds_per_second
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

func main() { __check(fmt.Sprint(time.Second.Nanoseconds()), "1000000000") }
