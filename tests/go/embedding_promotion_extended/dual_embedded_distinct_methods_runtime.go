// vybe-test: go/embedding_promotion_extended/dual_embedded_distinct_methods_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type north struct{}
func (north) letter() string { return "N" }
type east struct{}
func (east) letter() string { return "E" }
type compass struct { north
east }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { c := compass{}
__check(fmt.Sprint(c.north.letter()), "N")
__check(fmt.Sprint(c.east.letter()), "E") }
