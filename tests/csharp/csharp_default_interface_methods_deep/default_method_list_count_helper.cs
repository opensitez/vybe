// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_list_count_helper
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface ISize{int Len(System.Collections.Generic.List<int> xs)=>xs.Count;} class Measurer:ISize{} __P((new Measurer().Len(new System.Collections.Generic.List<int>{1,2,3})).ToString());
__Check("3");
