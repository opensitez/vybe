// vybe-test: go/net_textproto_compile/textproto_writer
// origin: languages/go/tests/go/test_net_textproto_compile.rs
// vybe-test-mode: compile

package main
import "net/textproto"
import "bytes"
func main() { w := textproto.NewWriter(bytes.NewBuffer(nil))
_ = w }
