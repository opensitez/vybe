// vybe-test: go/defer_panic_variants/named_return_pair_both_mutated_by_defers
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func stats() (total int, count int) { defer func() { count = 4 }()
defer func() { total = 9 }()
return 1, 2 }
func main() { t, c := stats()
fmt.Println(t)
fmt.Println(c)
}
