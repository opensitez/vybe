// vybe-test: go/time_location_zone/time_location_utc_singleton
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

func main() { t := time.Now().UTC()
__check(fmt.Sprint(t.Location() == time.UTC), "true") }
