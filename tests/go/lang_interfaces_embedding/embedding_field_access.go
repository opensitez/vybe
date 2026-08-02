// vybe-test: go/lang_interfaces_embedding/embedding_field_access
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type Inner struct { N int }
type Outer struct { Inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Outer{Inner: Inner{N: 4}}.N), "4") }
