// vybe-test: go/cover_net_extra/rpc_register_package_level
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/rpc"
type Args struct { A, B int }
type Arith int
func (t *Arith) Sub(args *Args, reply *int) error { return nil }
func main() { rpc.Register(new(Arith)) }
