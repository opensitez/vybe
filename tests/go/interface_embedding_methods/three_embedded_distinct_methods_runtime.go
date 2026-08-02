// vybe-test: go/interface_embedding_methods/three_embedded_distinct_methods_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type alpha interface { a() int }
type beta interface { b() int }
type gamma interface { c() int }
type combo interface { alpha
beta
gamma }
type triple struct{}
func (triple) a() int { return 1 }
func (triple) b() int { return 2 }
func (triple) c() int { return 3 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value combo = triple{}
__check(fmt.Sprint(value.a()), "1")
__check(fmt.Sprint(value.b()), "2")
__check(fmt.Sprint(value.c()), "3") }
