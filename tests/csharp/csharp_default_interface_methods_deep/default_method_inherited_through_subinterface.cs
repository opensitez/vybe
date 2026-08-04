// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_inherited_through_subinterface
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

interface IRoot{int Base()=>1;} interface IChild:IRoot{int Child()=>Base()+1;} class Node:IChild{} __P((new Node().Child()).ToString());
__Check("2");
