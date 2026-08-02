// vybe-test: go/label_break_continue/labeled_break_exits_outer_loop
// origin: languages/go/tests/go/test_label_break_continue.rs

package main
import "fmt"
func main() { sum := 0
outer: for i := 0; i < 3; i++ { for j := 0; j < 3; j++ { if i == 1 && j == 1 { break outer }
sum++ } }
fmt.Println(sum) }
