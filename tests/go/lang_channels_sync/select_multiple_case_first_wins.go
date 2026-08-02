// vybe-test: go/lang_channels_sync/select_multiple_case_first_wins
// origin: languages/go/tests/go/test_lang_channels_sync.rs

package main
import "fmt"
func main() { a := make(chan int, 1)
b := make(chan int, 1)
a <- 1
b <- 2
select { case v := <-a: fmt.Println(v)
case v := <-b: fmt.Println(v) } }
