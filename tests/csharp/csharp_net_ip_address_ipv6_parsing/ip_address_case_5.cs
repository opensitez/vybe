// vybe-test: csharp/csharp_net_ip_address_ipv6_parsing/ip_address_case_5

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var ip = System.Net.IPAddress.Parse("::1");
__P(ip.AddressFamily.ToString());
__Check("InterNetworkV6");
