// vybe-test: csharp/csharp_oop_file_local_types_visibility/file_local_type_case_12

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

__P(LocalHelper_12.Value.ToString());
__Check("12");

file class LocalHelper_12 {
    public static int Value => 12;
}
