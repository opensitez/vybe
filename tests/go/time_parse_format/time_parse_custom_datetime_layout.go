// vybe-test: go/time_parse_format/time_parse_custom_datetime_layout
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

func main() { t, _ := time.Parse("2006-01-02 15:04:05", "2020-01-02 03:04:05")
__check(fmt.Sprint(t.Hour()), "3") }
