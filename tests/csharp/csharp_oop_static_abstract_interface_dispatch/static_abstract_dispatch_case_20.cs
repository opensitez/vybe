// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_20

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

string res = FactoryCaller_20.Call();
__P(res);
__Check("Instance_20");

interface IFactory_20<TSelf> where TSelf : IFactory_20<TSelf> {
    static abstract string Create();
}
class FactoryImpl_20 : IFactory_20<FactoryImpl_20> {
    public static string Create() => "Instance_20";
}
class FactoryCaller_20 {
    public static string Call() => FactoryImpl_20.Create();
}
