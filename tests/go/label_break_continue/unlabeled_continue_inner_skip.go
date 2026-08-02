// vybe-test: go/label_break_continue/unlabeled_continue_inner_skip
// origin: languages/go/tests/go/test_label_break_continue.rs

package main
import "fmt"
func main() { sum := 0
for i := 0; i < 4; i++ { if i == 2 { continue }
sum += i }
fmt.Println(sum) }
