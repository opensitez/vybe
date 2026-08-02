// vybe-test: go/channel_close_range/close_channel_range_sum
// origin: languages/go/tests/go/test_channel_close_range.rs

package main
import "fmt"
func main() { ch := make(chan int, 3)
ch <- 1
ch <- 2
ch <- 3
close(ch)
sum := 0
for v := range ch { sum += v }
fmt.Println(sum) }
