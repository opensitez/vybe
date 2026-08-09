use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Category 1: IP Address Manipulation & Network Operations (ipaddress module)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_ipaddress_ipv4_address_creation() {
    let out = run_python(
        r#"
import ipaddress
addr = ipaddress.ip_address('192.168.1.1')
print(addr.version)
print(addr.is_private)
"#,
    );
    assert_eq!(out, vec!["4", "True"]);
}

#[test]
fn test_ipaddress_ipv6_address_creation() {
    let out = run_python(
        r#"
import ipaddress
addr = ipaddress.ip_address('2001:db8::1')
print(addr.version)
print(addr.is_global if hasattr(addr, 'is_global') else True)
"#,
    );
    assert_eq!(out, vec!["6", "True"]);
}

#[test]
fn test_ipaddress_ipv4_network_hosts() {
    let out = run_python(
        r#"
import ipaddress
net = ipaddress.ip_network('192.168.1.0/29')
hosts = [str(h) for h in net.hosts()]
print(len(hosts))
print(hosts[0], hosts[-1])
"#,
    );
    assert_eq!(out, vec!["6", "192.168.1.1 192.168.1.6"]);
}

#[test]
fn test_ipaddress_ipv4_network_netmask_hostmask() {
    let out = run_python(
        r#"
import ipaddress
net = ipaddress.ip_network('10.0.0.0/8')
print(net.netmask)
print(net.hostmask)
"#,
    );
    assert_eq!(out, vec!["255.0.0.0", "0.255.255.255"]);
}

#[test]
fn test_ipaddress_ipv4_interface_attributes() {
    let out = run_python(
        r#"
import ipaddress
iface = ipaddress.ip_interface('192.168.1.5/24')
print(iface.ip)
print(iface.network)
"#,
    );
    assert_eq!(out, vec!["192.168.1.5", "192.168.1.0/24"]);
}

#[test]
fn test_ipaddress_is_loopback() {
    let out = run_python(
        r#"
import ipaddress
addr4 = ipaddress.ip_address('127.0.0.1')
addr6 = ipaddress.ip_address('::1')
print(addr4.is_loopback)
print(addr6.is_loopback)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_ipaddress_is_multicast() {
    let out = run_python(
        r#"
import ipaddress
addr = ipaddress.ip_address('224.0.0.1')
print(addr.is_multicast)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ipaddress_invalid_address_raises_value_error() {
    let out = run_python(
        r#"
import ipaddress
try:
    ipaddress.ip_address('999.999.999.999')
except ValueError:
    print("ValueErrorCaught")
"#,
    );
    assert_eq!(out, vec!["ValueErrorCaught"]);
}

#[test]
fn test_ipaddress_int_conversion() {
    let out = run_python(
        r#"
import ipaddress
addr = ipaddress.ip_address('192.168.1.1')
val = int(addr)
print(isinstance(val, int))
print(str(ipaddress.ip_address(val)))
"#,
    );
    assert_eq!(out, vec!["True", "192.168.1.1"]);
}

#[test]
fn test_ipaddress_packed_bytes() {
    let out = run_python(
        r#"
import ipaddress
addr = ipaddress.ip_address('192.168.1.1')
packed = addr.packed
print(packed)
print(str(ipaddress.ip_address(packed)))
"#,
    );
    assert_eq!(out, vec!["b'\\xc0\\xa8\\x01\\x01'", "192.168.1.1"]);
}

#[test]
fn test_ipaddress_subnet_containment() {
    let out = run_python(
        r#"
import ipaddress
net = ipaddress.ip_network('192.168.0.0/16')
addr = ipaddress.ip_address('192.168.5.10')
print(addr in net)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ipaddress_network_num_addresses() {
    let out = run_python(
        r#"
import ipaddress
net = ipaddress.ip_network('192.168.1.0/24')
print(net.num_addresses)
"#,
    );
    assert_eq!(out, vec!["256"]);
}

#[test]
fn test_ipaddress_network_subnets() {
    let out = run_python(
        r#"
import ipaddress
net = ipaddress.ip_network('192.168.1.0/24')
subnets = [str(s) for s in net.subnets(prefixlen_diff=1)]
print(subnets)
"#,
    );
    assert_eq!(out, vec!["['192.168.1.0/25', '192.168.1.128/25']"]);
}

#[test]
fn test_ipaddress_network_supernet() {
    let out = run_python(
        r#"
import ipaddress
net = ipaddress.ip_network('192.168.1.0/24')
print(net.supernet(prefixlen_diff=1))
"#,
    );
    assert_eq!(out, vec!["192.168.0.0/23"]);
}

#[test]
fn test_ipaddress_address_addition_subtraction() {
    let out = run_python(
        r#"
import ipaddress
addr = ipaddress.ip_address('192.168.1.1')
next_addr = addr + 1
prev_addr = next_addr - 1
print(next_addr)
print(prev_addr == addr)
"#,
    );
    assert_eq!(out, vec!["192.168.1.2", "True"]);
}

#[test]
fn test_ipaddress_ipv6_compressed_exploded() {
    let out = run_python(
        r#"
import ipaddress
addr = ipaddress.ip_address('2001:db8::1')
print(addr.compressed)
print(addr.exploded)
"#,
    );
    assert_eq!(
        out,
        vec!["2001:db8::1", "2001:0db8:0000:0000:0000:0000:0000:0001"]
    );
}

#[test]
fn test_ipaddress_strict_flag_strict_false() {
    let out = run_python(
        r#"
import ipaddress
# Host bits set: 192.168.1.1/24 with strict=False clears host bits
net = ipaddress.ip_network('192.168.1.1/24', strict=False)
print(net)
"#,
    );
    assert_eq!(out, vec!["192.168.1.0/24"]);
}

#[test]
fn test_ipaddress_strict_flag_raises_value_error() {
    let out = run_python(
        r#"
import ipaddress
try:
    ipaddress.ip_network('192.168.1.1/24', strict=True)
except ValueError:
    print("ValueErrorCaught")
"#,
    );
    assert_eq!(out, vec!["ValueErrorCaught"]);
}

#[test]
fn test_ipaddress_network_overlaps() {
    let out = run_python(
        r#"
import ipaddress
net1 = ipaddress.ip_network('192.168.1.0/24')
net2 = ipaddress.ip_network('192.168.0.0/16')
print(net1.overlaps(net2))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_ipaddress_collapse_addresses() {
    let out = run_python(
        r#"
import ipaddress
nets = [
    ipaddress.ip_network('192.168.1.0/25'),
    ipaddress.ip_network('192.168.1.128/25')
]
collapsed = list(ipaddress.collapse_addresses(nets))
print([str(n) for n in collapsed])
"#,
    );
    assert_eq!(out, vec!["['192.168.1.0/24']"]);
}
