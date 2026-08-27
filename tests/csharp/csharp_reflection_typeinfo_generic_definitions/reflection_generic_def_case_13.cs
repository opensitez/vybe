// vybe-test: csharp/csharp_reflection_typeinfo_generic_definitions/reflection_generic_def_case_13

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

Type t = typeof(System.Collections.Generic.List<>);
__P(t.IsGenericTypeDefinition.ToString());
Type constructed = t.MakeGenericType(typeof(int));
__P(constructed.GenericTypeArguments[0].Name);
__Check("True\nInt32");
