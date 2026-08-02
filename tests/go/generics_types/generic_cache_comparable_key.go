// vybe-test: go/generics_types/generic_cache_comparable_key
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Cache[K comparable, V any] struct { m map[K]V }
func (c *Cache[K, V]) Put(k K, v V) { if c.m == nil { c.m = make(map[K]V) }
c.m[k] = v }
func (c Cache[K, V]) Get(k K) (V, bool) { v, ok := c.m[k]
return v, ok }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var c Cache[string, int]
c.Put("x", 3)
v, ok := c.Get("x")
__check(fmt.Sprint(ok), "true")
__check(fmt.Sprint(v), "3") }
