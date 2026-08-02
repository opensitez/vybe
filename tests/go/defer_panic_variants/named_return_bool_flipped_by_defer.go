// vybe-test: go/defer_panic_variants/named_return_bool_flipped_by_defer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func check() (ok bool) { defer func() { ok = true }()
return false }
func main() { fmt.Println(check())
}
