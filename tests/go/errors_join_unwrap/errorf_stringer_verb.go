// vybe-test: go/errors_join_unwrap/errorf_stringer_verb
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
type id int
func (i id) String() string { return fmt.Sprintf("ID-%d", i) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("entity %s", id(5))
__check(fmt.Sprint(err.Error()), "entity ID-5") }
