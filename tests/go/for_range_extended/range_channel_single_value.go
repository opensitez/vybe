// vybe-test: go/for_range_extended/range_channel_single_value
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { ch := make(chan int, 1)
ch <- 99
close(ch)
last := 0
for v := range ch { last = v }
fmt.Println(last) }
