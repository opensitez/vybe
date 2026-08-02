// vybe-test: go/init_blank_import/init_fills_map_lookup_before_main
// origin: languages/go/tests/go/test_init_blank_import.rs

package main
import "fmt"
var table = map[string]int{}
func init() { table["go"] = 7 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(table["go"]), "7") }
