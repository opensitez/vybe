// vybe-test: go/time_package/time_parse_rfc3339
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

func main() { t, _ := time.Parse(time.RFC3339, "2021-06-15T12:00:00Z")
__check(fmt.Sprint(t.Month()), "June") }
