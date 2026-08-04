// vybe-test: go/net_netip/netip_addr_is4_on_ipv4_literal
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

func main() { a, _ := netip.ParseAddr("8.8.8.8")
__p(fmt.Sprint(a.Is4())) 
__check("true")
}
