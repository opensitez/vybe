// vybe-test: go/label_break_continue/labeled_break_on_search_found
// origin: languages/go/tests/go/test_label_break_continue.rs

package main
import "fmt"
func main() { grid := [][]int{{1,2},{3,4}}
found := -1
search: for r := 0; r < len(grid); r++ { for c := 0; c < len(grid[r]); c++ { if grid[r][c] == 3 { found = r*10 + c
break search } } }
fmt.Println(found) }
