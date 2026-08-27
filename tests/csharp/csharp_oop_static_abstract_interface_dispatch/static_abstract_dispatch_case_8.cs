// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_8

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

string res = FactoryCaller_8.Call();
__P(res);
__Check("Instance_8");

interface IFactory_8<TSelf> where TSelf : IFactory_8<TSelf> {
    static abstract string Create();
}
class FactoryImpl_8 : IFactory_8<FactoryImpl_8> {
    public static string Create() => "Instance_8";
}
class FactoryCaller_8 {
    public static string Call() => FactoryImpl_8.Create();
}
