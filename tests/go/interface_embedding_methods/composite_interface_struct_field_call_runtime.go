// vybe-test: go/interface_embedding_methods/composite_interface_struct_field_call_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type runner interface { run() int }
type athlete interface { runner }
type team struct { lead athlete }
type sprinter struct { pace int }
func (s sprinter) run() int { return s.pace }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { squad := team{lead: sprinter{pace: 42}}
__check(fmt.Sprint(squad.lead.run()), "42") }
