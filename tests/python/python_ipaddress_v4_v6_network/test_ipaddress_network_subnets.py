# vybe-test: python/python_ipaddress_v4_v6_network/test_ipaddress_network_subnets
# origin: languages/python/tests/python/test_python_ipaddress_v4_v6_network.rs

import ipaddress
net = ipaddress.ip_network('192.168.1.0/24')
subnets = [str(s) for s in net.subnets(prefixlen_diff=1)]
print(subnets)
