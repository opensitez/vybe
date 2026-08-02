// vybe-test: go/defer_panic_variants/defer_inner_frame_finishes_before_outer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func inner() { defer fmt.Println("inner")
}
func main() { defer fmt.Println("outer")
inner()
}
