// vybe-test: go/net_netip/netip_prefix_bits_returns_mask_length
// origin: languages/go/tests/go/test_net_netip.rs

package main
import "fmt"
import "net/netip"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { p, _ := netip.ParsePrefix("192.0.2.0/26")
__p(fmt.Sprint(p.Bits())) 
__check("26")
}
