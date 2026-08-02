// vybe-test: go/net_textproto_compile/textproto_mime_header
// origin: languages/go/tests/go/test_net_textproto_compile.rs
// vybe-test-mode: compile

package main
import "net/textproto"
func main() { h := make(textproto.MIMEHeader)
h.Set("K", "V") }
