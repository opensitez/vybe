// vybe-test: csharp/csharp_numerics_int128_uint128/int128_in_generic_list

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

var list = new System.Collections.Generic.List<Int128>();
list.Add(Int128.Parse("500000000000000000"));
__P(list.Count.ToString());
__P(list[0].ToString());
__Check("1\n500000000000000000");
