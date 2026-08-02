// vybe-test: go/time_location_zone/time_truncate_sub_hour
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

func main() { t := time.Date(2023, 6, 1, 10, 30, 0, 0, time.UTC)
truncated := t.Truncate(30 * time.Minute)
__check(fmt.Sprint(truncated.Minute()), "30") }
