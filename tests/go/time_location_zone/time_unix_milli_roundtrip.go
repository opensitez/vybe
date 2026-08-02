// vybe-test: go/time_location_zone/time_unix_milli_roundtrip
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

func main() { ms := int64(1609459200123)
t := time.UnixMilli(ms)
__check(fmt.Sprint(t.UnixMilli()), "1609459200123") }
