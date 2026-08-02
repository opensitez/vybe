// vybe-test: go/time_package/time_format_rfc3339_utc
// origin: languages/go/tests/go/test_time_package.rs

package main
import "fmt"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := time.Date(2020, 1, 2, 3, 4, 5, 0, time.UTC)
__check(fmt.Sprint(t.Format(time.RFC3339)), "2020-01-02T03:04:05Z") }
