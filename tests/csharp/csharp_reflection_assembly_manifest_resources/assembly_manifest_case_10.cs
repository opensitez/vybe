// vybe-test: csharp/csharp_reflection_assembly_manifest_resources/assembly_manifest_case_10

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

var asm = typeof(object).Assembly;
__P((asm.GetName().Name).ToString());
__Check("System.Private.CoreLib");
