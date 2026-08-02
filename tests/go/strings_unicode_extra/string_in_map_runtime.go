// vybe-test: go/strings_unicode_extra/string_in_map_runtime
// origin: languages/go/tests/go/test_strings_unicode_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]string{"lang": "go"}
__check(fmt.Sprint(values["lang"]), "go")
}
