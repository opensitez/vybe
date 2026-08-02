// vybe-test: go/time_parse_format/time_format_kitchen_pm
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

func main() { t := time.Date(2020, 3, 15, 14, 30, 0, 0, time.UTC)
__check(fmt.Sprint(t.Format(time.Kitchen)), "2:30PM") }
