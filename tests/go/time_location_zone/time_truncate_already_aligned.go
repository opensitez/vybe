// vybe-test: go/time_location_zone/time_truncate_already_aligned
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

func main() { t := time.Date(2023, 1, 1, 10, 0, 0, 0, time.UTC)
truncated := t.Truncate(time.Hour)
__check(fmt.Sprint(truncated.Equal(t)), "true") }
