// vybe-test: go/variadic_advanced/forward_variadic_to_peer_direct
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func sink(nums ...int) int { t := 0
for _, n := range nums { t += n }
return t }
func relay(nums ...int) int { return sink(nums...) }
func main() { fmt.Println(relay(1, 2, 3)) }
