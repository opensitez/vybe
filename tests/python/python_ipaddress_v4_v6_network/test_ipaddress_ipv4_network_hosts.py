# vybe-test: python/python_ipaddress_v4_v6_network/test_ipaddress_ipv4_network_hosts
# origin: languages/python/tests/python/test_python_ipaddress_v4_v6_network.rs

import ipaddress
net = ipaddress.ip_network('192.168.1.0/29')
hosts = [str(h) for h in net.hosts()]
print(len(hosts))
print(hosts[0], hosts[-1])
