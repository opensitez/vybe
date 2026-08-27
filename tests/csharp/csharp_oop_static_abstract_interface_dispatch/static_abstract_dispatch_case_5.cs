// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_5

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

string res = FactoryCaller_5.Call();
__P(res);
__Check("Instance_5");

interface IFactory_5<TSelf> where TSelf : IFactory_5<TSelf> {
    static abstract string Create();
}
class FactoryImpl_5 : IFactory_5<FactoryImpl_5> {
    public static string Create() => "Instance_5";
}
class FactoryCaller_5 {
    public static string Call() => FactoryImpl_5.Create();
}
