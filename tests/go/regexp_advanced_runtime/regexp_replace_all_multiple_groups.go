// vybe-test: go/regexp_advanced_runtime/regexp_replace_all_multiple_groups
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs

package main
import "fmt"
import "regexp"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { re := regexp.MustCompile(`(\d{2})-(\d{2})-(\d{4})`)
__check(fmt.Sprint(re.ReplaceAllString("06-30-2024", "$3/$1/$2")), "2024/06/30") }
