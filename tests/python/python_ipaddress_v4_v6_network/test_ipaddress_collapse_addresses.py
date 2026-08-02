# vybe-test: python/python_ipaddress_v4_v6_network/test_ipaddress_collapse_addresses
# origin: languages/python/tests/python/test_python_ipaddress_v4_v6_network.rs

import ipaddress
nets = [
    ipaddress.ip_network('192.168.1.0/25'),
    ipaddress.ip_network('192.168.1.128/25')
]
collapsed = list(ipaddress.collapse_addresses(nets))
print([str(n) for n in collapsed])
