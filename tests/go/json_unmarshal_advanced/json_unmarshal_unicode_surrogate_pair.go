// vybe-test: go/json_unmarshal_advanced/json_unmarshal_unicode_surrogate_pair
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s string
json.Unmarshal([]byte(`"\uD83D\uDE00"`), &s)
__check(fmt.Sprint(len([]rune(s))), "1") }
