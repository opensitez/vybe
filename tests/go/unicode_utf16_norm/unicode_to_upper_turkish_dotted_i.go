// vybe-test: go/unicode_utf16_norm/unicode_to_upper_turkish_dotted_i
// origin: languages/go/tests/go/test_unicode_utf16_norm.rs

package main
import "fmt"
import "unicode"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(int(unicode.ToUpper('i'))), "73") }
