// vybe-test: go/net_textproto_compile/net_dial_tcp
// origin: languages/go/tests/go/test_net_textproto_compile.rs
// vybe-test-mode: compile

package main
import "net"
func main() { c, _ := net.Dial("tcp", "127.0.0.1:9")
if c != nil { c.Close() } }
