// vybe-test: go/time_location_zone/time_unix_micro_roundtrip
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

func main() { us := int64(1609459200456789)
t := time.UnixMicro(us)
__check(fmt.Sprint(t.UnixMicro()), "1609459200456789") }
