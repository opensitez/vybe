// vybe-test: go/embedding_promotion_extended/two_level_embedded_interface_field_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type speaker interface { talk() string }
type bot struct{}
func (bot) talk() string { return "beep" }
type host struct { speaker }
type rack struct { host }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := rack{host: host{speaker: bot{}}}
__check(fmt.Sprint(r.talk()), "beep") }
