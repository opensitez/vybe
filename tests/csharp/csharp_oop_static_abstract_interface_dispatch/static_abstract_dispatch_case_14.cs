// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_14

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

string res = FactoryCaller_14.Call();
__P(res);
__Check("Instance_14");

interface IFactory_14<TSelf> where TSelf : IFactory_14<TSelf> {
    static abstract string Create();
}
class FactoryImpl_14 : IFactory_14<FactoryImpl_14> {
    public static string Create() => "Instance_14";
}
class FactoryCaller_14 {
    public static string Call() => FactoryImpl_14.Create();
}
