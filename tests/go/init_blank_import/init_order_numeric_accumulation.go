// vybe-test: go/init_blank_import/init_order_numeric_accumulation
// origin: languages/go/tests/go/test_init_blank_import.rs

package main
import "fmt"
var total int
func init() { total += 1 }
func init() { total += 10 }
func init() { total += 100 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(total), "111") }
