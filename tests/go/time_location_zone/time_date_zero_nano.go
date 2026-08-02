// vybe-test: go/time_location_zone/time_date_zero_nano
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

func main() { t := time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC)
__check(fmt.Sprint(t.Nanosecond()), "0")
__check(fmt.Sprint(t.Second()), "0") }
