// vybe-test: go/time_parse_format/time_parse_rfc822_utc
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

func main() { t, _ := time.Parse(time.RFC822, "02 Jan 20 03:04 UTC")
__check(fmt.Sprint(t.Day()), "2") }
