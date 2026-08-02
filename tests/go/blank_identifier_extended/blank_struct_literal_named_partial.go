// vybe-test: go/blank_identifier_extended/blank_struct_literal_named_partial
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
type cfg struct { host string
port int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := cfg{port: 8080}
__check(fmt.Sprint(c.port), "8080") }
