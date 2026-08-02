// vybe-test: go/for_range_extended/range_channel_buffered_sum
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { ch := make(chan int, 3)
ch <- 2
ch <- 4
ch <- 6
close(ch)
total := 0
for v := range ch { total += v }
fmt.Println(total) }
