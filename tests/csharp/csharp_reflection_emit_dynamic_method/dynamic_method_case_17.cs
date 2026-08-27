// vybe-test: csharp/csharp_reflection_emit_dynamic_method/dynamic_method_case_17

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

var op = System.Reflection.Emit.OpCodes.Dup;
__P(op.Name);
__P(((int)op.Value).ToString());
__Check("dup\n37");
