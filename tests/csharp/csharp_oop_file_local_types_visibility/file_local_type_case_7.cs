// vybe-test: csharp/csharp_oop_file_local_types_visibility/file_local_type_case_7

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

__P(LocalHelper_7.Value.ToString());
__Check("7");

file class LocalHelper_7 {
    public static int Value => 7;
}
