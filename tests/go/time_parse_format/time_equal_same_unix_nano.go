// vybe-test: go/time_parse_format/time_equal_same_unix_nano
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

func main() { a := time.Unix(5, 100)
b := time.Unix(5, 100)
__check(fmt.Sprint(a.Equal(b)), "true") }
