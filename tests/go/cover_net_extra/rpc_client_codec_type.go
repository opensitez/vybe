// vybe-test: go/cover_net_extra/rpc_client_codec_type
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
type ClientCodec = rpc.ClientCodec
func main() { var cc ClientCodec
_ = cc }
