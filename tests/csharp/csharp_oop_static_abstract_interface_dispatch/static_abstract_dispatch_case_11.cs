// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_11

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

string res = FactoryCaller_11.Call();
__P(res);
__Check("Instance_11");

interface IFactory_11<TSelf> where TSelf : IFactory_11<TSelf> {
    static abstract string Create();
}
class FactoryImpl_11 : IFactory_11<FactoryImpl_11> {
    public static string Create() => "Instance_11";
}
class FactoryCaller_11 {
    public static string Call() => FactoryImpl_11.Create();
}
