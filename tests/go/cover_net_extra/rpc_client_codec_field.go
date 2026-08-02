// vybe-test: go/cover_net_extra/rpc_client_codec_field
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
func main() { c, _ := rpc.Dial("tcp", "127.0.0.1:9999")
if c != nil { defer c.Close()
_ = c.Codec } }
