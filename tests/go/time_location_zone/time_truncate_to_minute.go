// vybe-test: go/time_location_zone/time_truncate_to_minute
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

func main() { t := time.Date(2023, 1, 1, 10, 15, 45, 0, time.UTC)
truncated := t.Truncate(time.Minute)
__check(fmt.Sprint(truncated.Second()), "0") }
