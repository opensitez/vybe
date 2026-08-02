// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_date_iso
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

func main() { re := regexp.MustCompile(`(\d{4})-(\d{2})-(\d{2})`)
m := re.FindAllStringSubmatch("2024-06-30", -1)
__check(fmt.Sprint(m[0][1]), "2024")
__check(fmt.Sprint(m[0][3]), "30") }
