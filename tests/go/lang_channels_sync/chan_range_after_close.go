// vybe-test: go/lang_channels_sync/chan_range_after_close
// origin: languages/go/tests/go/test_lang_channels_sync.rs

package main
import "fmt"
func main() { ch := make(chan int, 2)
ch <- 1
ch <- 2
close(ch)
n := 0
for range ch { n++ }
fmt.Println(n) }
