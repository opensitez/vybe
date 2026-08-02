// vybe-test: go/time_parse_format/time_parse_duration_hours_minutes
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

func main() { d, _ := time.ParseDuration("2h30m")
__check(fmt.Sprint(d.Minutes()), "150") }
