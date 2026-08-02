// vybe-test: go/lang_channels_sync/select_default_nonblocking
// origin: languages/go/tests/go/test_lang_channels_sync.rs

package main
import "fmt"
func main() { ch := make(chan int)
select { case <-ch: fmt.Println("recv")
default: fmt.Println("def") } }
