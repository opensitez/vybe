//! net and net/textproto compile patterns.


go_compile_cases! {
    net_lookup_host_compile => "package main; import \"net\"; func main() { _, _ = net.LookupHost(\"localhost\") }",
    net_resolve_tcp_addr => "package main; import \"net\"; func main() { _, _ = net.ResolveTCPAddr(\"tcp\", \":80\") }",
    net_resolve_udp_addr => "package main; import \"net\"; func main() { _, _ = net.ResolveUDPAddr(\"udp\", \":53\") }",
    net_listen_tcp => "package main; import \"net\"; func main() { ln, _ := net.Listen(\"tcp\", \":0\"); if ln != nil { ln.Close() } }",
    net_dial_tcp => "package main; import \"net\"; func main() { c, _ := net.Dial(\"tcp\", \"127.0.0.1:9\"); if c != nil { c.Close() } }",
    net_join_host_port => "package main; import \"net\"; func main() { _, _ = net.JoinHostPort(\"127.0.0.1\", \"80\") }",
    net_split_host_port => "package main; import \"net\"; func main() { _, _, _ = net.SplitHostPort(\"127.0.0.1:80\") }",
    net_cidr_mask => "package main; import \"net\"; func main() { _ = net.CIDRMask(24, 32) }",
    net_parse_cidr => "package main; import \"net\"; func main() { _, _, _ = net.ParseCIDR(\"10.0.0.0/8\") }",
    net_ip_is_loopback => "package main; import \"net\"; func main() { ip := net.ParseIP(\"127.0.0.1\"); _ = ip.IsLoopback() }",
    textproto_reader => "package main; import \"net/textproto\"; import \"strings\"; func main() { r := textproto.NewReader(strings.NewReader(\"\")); _ = r }",
    textproto_writer => "package main; import \"net/textproto\"; import \"bytes\"; func main() { w := textproto.NewWriter(bytes.NewBuffer(nil)); _ = w }",
    textproto_mime_header => "package main; import \"net/textproto\"; func main() { h := make(textproto.MIMEHeader); h.Set(\"K\", \"V\") }",
}
