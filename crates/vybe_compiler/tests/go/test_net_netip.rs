//! net/netip: ParseAddr, IPv4/IPv6, Is4/Is6, String roundtrip, Prefix, Contains, Mask, AddrPort.

go_run_cases! {
    netip_parse_ipv4_dotted_string => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"192.168.0.1\"); fmt.Println(a.String()) }",
        vec!["192.168.0.1"]
    ),
    netip_parse_ipv4_loopback => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"127.0.0.1\"); fmt.Println(a.String()) }",
        vec!["127.0.0.1"]
    ),
    netip_parse_ipv6_full_form => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"2001:db8::1\"); fmt.Println(a.String()) }",
        vec!["2001:db8::1"]
    ),
    netip_parse_ipv6_loopback_shorthand => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"::1\"); fmt.Println(a.String()) }",
        vec!["::1"]
    ),
    netip_ipv4_constructor_string => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { fmt.Println(netip.IPv4(10, 0, 0, 1).String()) }",
        vec!["10.0.0.1"]
    ),
    netip_ipv4_all_ones_broadcast => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { fmt.Println(netip.IPv4(255, 255, 255, 255).String()) }",
        vec!["255.255.255.255"]
    ),
    netip_addr_is4_on_ipv4_literal => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"8.8.8.8\"); fmt.Println(a.Is4()) }",
        vec!["true"]
    ),
    netip_addr_is4_false_on_ipv6 => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"::1\"); fmt.Println(a.Is4()) }",
        vec!["false"]
    ),
    netip_addr_is6_on_ipv6_literal => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"fe80::1\"); fmt.Println(a.Is6()) }",
        vec!["true"]
    ),
    netip_addr_is6_false_on_ipv4 => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"1.2.3.4\"); fmt.Println(a.Is6()) }",
        vec!["false"]
    ),
    netip_string_roundtrip_ipv4 => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { orig := \"203.0.113.5\"; a, _ := netip.ParseAddr(orig); fmt.Println(a.String() == orig) }",
        vec!["true"]
    ),
    netip_string_roundtrip_ipv6 => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { orig := \"2001:db8:85a3::8a2e:370:7334\"; a, _ := netip.ParseAddr(orig); fmt.Println(a.String() == orig) }",
        vec!["true"]
    ),
    netip_prefix_parse_slash_24 => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { p, _ := netip.ParsePrefix(\"192.168.1.0/24\"); fmt.Println(p.String()) }",
        vec!["192.168.1.0/24"]
    ),
    netip_prefix_parse_slash_32_host => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { p, _ := netip.ParsePrefix(\"10.0.0.5/32\"); fmt.Println(p.String()) }",
        vec!["10.0.0.5/32"]
    ),
    netip_prefix_contains_addr_inside => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { p, _ := netip.ParsePrefix(\"10.0.0.0/8\"); a, _ := netip.ParseAddr(\"10.1.2.3\"); fmt.Println(p.Contains(a)) }",
        vec!["true"]
    ),
    netip_prefix_contains_addr_outside => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { p, _ := netip.ParsePrefix(\"10.0.0.0/8\"); a, _ := netip.ParseAddr(\"172.16.0.1\"); fmt.Println(p.Contains(a)) }",
        vec!["false"]
    ),
    netip_prefix_bits_returns_mask_length => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { p, _ := netip.ParsePrefix(\"192.0.2.0/26\"); fmt.Println(p.Bits()) }",
        vec!["26"]
    ),
    netip_prefix_masked_zeros_host_bits => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { p, _ := netip.ParsePrefix(\"192.168.1.129/24\"); fmt.Println(p.Masked().String()) }",
        vec!["192.168.1.0/24"]
    ),
    netip_addr_port_parse_host_port => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { ap, _ := netip.ParseAddrPort(\"127.0.0.1:8080\"); fmt.Println(ap.String()) }",
        vec!["127.0.0.1:8080"]
    ),
    netip_addr_port_port_field => (
        "package main; import \"fmt\"; import \"net/netip\"; func main() { ap, _ := netip.ParseAddrPort(\"[::1]:443\"); fmt.Println(ap.Port()) }",
        vec!["443"]
    ),
}

go_compile_cases! {
    netip_parse_addr_ipv4_mapped_ipv6 => "package main; import \"net/netip\"; func main() { _, _ = netip.ParseAddr(\"::ffff:192.0.2.1\") }",
    netip_parse_addr_unspecified_ipv4 => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"0.0.0.0\"); _ = a.IsUnspecified() }",
    netip_parse_addr_unspecified_ipv6 => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"::\"); _ = a.IsUnspecified() }",
    netip_ipv4_from_byte_slice => "package main; import \"net/netip\"; func main() { _, ok := netip.AddrFromSlice([]byte{1, 2, 3, 4}); _ = ok }",
    netip_ipv6_from_sixteen_byte_slice => "package main; import \"net/netip\"; func main() { _, ok := netip.AddrFromSlice(make([]byte, 16)); _ = ok }",
    netip_ipv4_constructor_zero_address => "package main; import \"net/netip\"; func main() { _ = netip.IPv4(0, 0, 0, 0).IsUnspecified() }",
    netip_addr_is_loopback_v4 => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"127.0.0.1\"); _ = a.IsLoopback() }",
    netip_addr_is_loopback_v6 => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"::1\"); _ = a.IsLoopback() }",
    netip_addr_is_private_rfc1918 => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"10.255.0.1\"); _ = a.IsPrivate() }",
    netip_addr_is_global_unicast => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"8.8.8.8\"); _ = a.IsGlobalUnicast() }",
    netip_addr_is_link_local_unicast => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"fe80::1\"); _ = a.IsLinkLocalUnicast() }",
    netip_addr_is_multicast_v4 => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"224.0.0.1\"); _ = a.IsMulticast() }",
    netip_addr_is_multicast_v6 => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"ff02::1\"); _ = a.IsMulticast() }",
    netip_addr_unmap_ipv4_mapped => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"::ffff:192.0.2.1\"); _ = a.Unmap().String() }",
    netip_addr_with_zone_id => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"fe80::1%eth0\"); _ = a.WithZone(\"eth0\").Zone() }",
    netip_addr_compare_equal => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"1.1.1.1\"); b, _ := netip.ParseAddr(\"1.1.1.1\"); _ = a.Compare(b) }",
    netip_addr_equal_method => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"1.1.1.1\"); b, _ := netip.ParseAddr(\"1.1.1.1\"); _ = a.Equal(b) }",
    netip_addr_less_than_ordering => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"1.0.0.1\"); b, _ := netip.ParseAddr(\"1.0.0.2\"); _ = a.Less(b) }",
    netip_prefix_from_addr_and_bits => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"192.168.0.0\"); _, _ = netip.PrefixFrom(a, 16) }",
    netip_prefix_ipv6_slash_64 => "package main; import \"net/netip\"; func main() { _, _ = netip.ParsePrefix(\"2001:db8::/64\") }",
    netip_prefix_is_valid_true => "package main; import \"net/netip\"; func main() { p, _ := netip.ParsePrefix(\"10.0.0.0/8\"); _ = p.IsValid() }",
    netip_prefix_addr_field => "package main; import \"net/netip\"; func main() { p, _ := netip.ParsePrefix(\"172.16.0.0/12\"); _ = p.Addr().String() }",
    netip_prefix_overlaps_shared_subnet => "package main; import \"net/netip\"; func main() { a, _ := netip.ParsePrefix(\"10.0.0.0/8\"); b, _ := netip.ParsePrefix(\"10.1.0.0/16\"); _ = a.Overlaps(b) }",
    netip_prefix_overlaps_disjoint => "package main; import \"net/netip\"; func main() { a, _ := netip.ParsePrefix(\"10.0.0.0/8\"); b, _ := netip.ParsePrefix(\"192.168.0.0/16\"); _ = a.Overlaps(b) }",
    netip_prefix_contains_prefix => "package main; import \"net/netip\"; func main() { outer, _ := netip.ParsePrefix(\"10.0.0.0/8\"); inner, _ := netip.ParsePrefix(\"10.1.0.0/16\"); _ = outer.ContainsPrefix(inner) }",
    netip_must_parse_addr_panics_on_bad => "package main; import \"net/netip\"; func main() { defer func() { _ = recover() }(); _ = netip.MustParseAddr(\"1.2.3.4\") }",
    netip_must_parse_prefix_valid => "package main; import \"net/netip\"; func main() { _ = netip.MustParsePrefix(\"0.0.0.0/0\").Bits() }",
    netip_addr_port_from_components => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"192.0.2.1\"); _ = netip.AddrPortFrom(a, 53).String() }",
    netip_addr_port_addr_field => "package main; import \"net/netip\"; func main() { ap, _ := netip.ParseAddrPort(\"203.0.113.7:9000\"); _ = ap.Addr().String() }",
    netip_addr_port_bracketed_ipv6 => "package main; import \"net/netip\"; func main() { _, _ = netip.ParseAddrPort(\"[2001:db8::1]:80\") }",
    netip_addr_as_slice_four_bytes => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"1.2.3.4\"); _ = len(a.AsSlice()) }",
    netip_addr_as_16_byte_slice_v6 => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"::1\"); _ = len(a.As16()) }",
    netip_addr_next_prev_adjacent => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"192.0.2.1\"); _ = a.Next().String(); _ = a.Prev().String() }",
    netip_prefix_from_ipv6_addr_thirty_two => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"2001:db8::\"); p, _ := netip.PrefixFrom(a, 32); _ = p.String() }",
    netip_parse_addr_invalid_returns_error => "package main; import \"net/netip\"; func main() { _, err := netip.ParseAddr(\"not-an-ip\"); _ = err != nil }",
    netip_parse_prefix_invalid_returns_error => "package main; import \"net/netip\"; func main() { _, err := netip.ParsePrefix(\"999.999.999.999/99\"); _ = err != nil }",
    netip_parse_addr_port_missing_port => "package main; import \"net/netip\"; func main() { _, err := netip.ParseAddrPort(\"127.0.0.1\"); _ = err != nil }",
    netip_addr_is4in6_mapped_form => "package main; import \"net/netip\"; func main() { a, _ := netip.ParseAddr(\"::ffff:1.2.3.4\"); _ = a.Is4In6() }",
    netip_prefix_string_roundtrip => "package main; import \"net/netip\"; func main() { s := \"198.51.100.0/24\"; p, _ := netip.ParsePrefix(s); _ = p.String() == s }",
    netip_addr_port_string_roundtrip => "package main; import \"net/netip\"; func main() { s := \"198.51.100.2:22\"; ap, _ := netip.ParseAddrPort(s); _ = ap.String() == s }",
}
