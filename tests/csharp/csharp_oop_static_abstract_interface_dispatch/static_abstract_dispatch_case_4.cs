// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_4

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

string res = FactoryCaller_4.Call();
__P(res);
__Check("Instance_4");

interface IFactory_4<TSelf> where TSelf : IFactory_4<TSelf> {
    static abstract string Create();
}
class FactoryImpl_4 : IFactory_4<FactoryImpl_4> {
    public static string Create() => "Instance_4";
}
class FactoryCaller_4 {
    public static string Call() => FactoryImpl_4.Create();
}
