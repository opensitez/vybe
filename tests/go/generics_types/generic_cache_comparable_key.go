// vybe-test: go/generics_types/generic_cache_comparable_key
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Cache[K comparable, V any] struct { m map[K]V }
func (c *Cache[K, V]) Put(k K, v V) { if c.m == nil { c.m = make(map[K]V) }
c.m[k] = v }
func (c Cache[K, V]) Get(k K) (V, bool) { v, ok := c.m[k]
return v, ok }
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { var c Cache[string, int]
c.Put("x", 3)
v, ok := c.Get("x")
__p(fmt.Sprint(ok))
__p(fmt.Sprint(v)) 
__check("true\n3")
}
