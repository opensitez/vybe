// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_6

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

string res = FactoryCaller_6.Call();
__P(res);
__Check("Instance_6");

interface IFactory_6<TSelf> where TSelf : IFactory_6<TSelf> {
    static abstract string Create();
}
class FactoryImpl_6 : IFactory_6<FactoryImpl_6> {
    public static string Create() => "Instance_6";
}
class FactoryCaller_6 {
    public static string Call() => FactoryImpl_6.Create();
}
