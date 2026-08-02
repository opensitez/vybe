// vybe-test: go/json_unmarshal_advanced/json_unmarshal_embedded_slice
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs

package main
import "fmt"
import "encoding/json"
type Tags []string
type Post struct { Tags
Title string }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p Post
json.Unmarshal([]byte(`{"Title":"t","Tags":["a","b"]}`), &p)
__check(fmt.Sprint(len(p.Tags)), "2")
__check(fmt.Sprint(p.Tags[1]), "b") }
