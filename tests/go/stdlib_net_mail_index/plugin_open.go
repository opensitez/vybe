// vybe-test: go/stdlib_net_mail_index/plugin_open
// origin: languages/go/tests/go/test_stdlib_net_mail_index.rs
// vybe-test-mode: compile

package main
import "plugin"
func main() { _, _ = plugin.Open("plugin.so") }
