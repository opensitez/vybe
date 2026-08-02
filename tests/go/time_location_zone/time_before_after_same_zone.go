// vybe-test: go/time_location_zone/time_before_after_same_zone
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

func main() { early := time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC)
late := time.Date(2020, 1, 2, 0, 0, 0, 0, time.UTC)
__check(fmt.Sprint(early.Before(late)), "true")
__check(fmt.Sprint(late.After(early)), "true") }
