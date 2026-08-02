// vybe-test: go/variadic_spread/forward_variadic_to_fmt_println_wrapper
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func show(parts ...interface{}) { fmt.Println(parts...) }
func main() { show("vybe", 42)
}
