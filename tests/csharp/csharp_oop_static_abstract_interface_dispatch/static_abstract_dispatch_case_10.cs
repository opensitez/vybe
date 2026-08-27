// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_10

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

string res = FactoryCaller_10.Call();
__P(res);
__Check("Instance_10");

interface IFactory_10<TSelf> where TSelf : IFactory_10<TSelf> {
    static abstract string Create();
}
class FactoryImpl_10 : IFactory_10<FactoryImpl_10> {
    public static string Create() => "Instance_10";
}
class FactoryCaller_10 {
    public static string Call() => FactoryImpl_10.Create();
}
