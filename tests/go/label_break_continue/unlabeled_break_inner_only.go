// vybe-test: go/label_break_continue/unlabeled_break_inner_only
// origin: languages/go/tests/go/test_label_break_continue.rs

package main
import "fmt"
func main() { total := 0
for i := 0; i < 3; i++ { for j := 0; j < 3; j++ { if j == 1 { break }
total++ } }
fmt.Println(total) }
