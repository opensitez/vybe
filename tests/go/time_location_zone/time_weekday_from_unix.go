// vybe-test: go/time_location_zone/time_weekday_from_unix
// origin: languages/go/tests/go/test_time_location_zone.rs

package main
import "fmt"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := time.Unix(86400, 0).UTC()
__check(fmt.Sprint(t.Weekday()), "Friday") }
