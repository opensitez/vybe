// vybe-test: go/lang_declarations_types/comparable_array_as_map_key
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[[1]int]string{[1]int{1}:"a"}
__check(fmt.Sprint(m[[1]int{1}]), "a") }
