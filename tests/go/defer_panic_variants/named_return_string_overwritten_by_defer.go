// vybe-test: go/defer_panic_variants/named_return_string_overwritten_by_defer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func greet() (msg string) { defer func() { msg = "bye" }()
return "hi" }
func main() { fmt.Println(greet())
}
