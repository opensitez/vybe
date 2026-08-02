// vybe-test: go/label_break_continue/labeled_continue_skips_outer_increment
// origin: languages/go/tests/go/test_label_break_continue.rs

package main
import "fmt"
func main() { count := 0
outer: for i := 0; i < 3; i++ { for j := 0; j < 2; j++ { if j == 1 { continue outer }
count++ } }
fmt.Println(count) }
