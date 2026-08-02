// vybe-test: go/embedding_promotion_extended/embedded_map_field_lookup_promoted_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type inner struct { data map[string]int }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{data: map[string]int{"k": 4}}}
__check(fmt.Sprint(o.data["k"]), "4") }
