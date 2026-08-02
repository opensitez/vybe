// vybe-test: go/cover_net_extra/rpc_server_codec_type
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
type ServerCodec = rpc.ServerCodec
func main() { var sc ServerCodec
_ = sc }
