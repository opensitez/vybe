// vybe-test: go/time_parse_format/time_parse_custom_date_layout
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

func main() { t, _ := time.Parse("2006-01-02", "2019-07-04")
__check(fmt.Sprint(t.Year()), "2019") }
