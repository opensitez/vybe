// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_7

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

string res = FactoryCaller_7.Call();
__P(res);
__Check("Instance_7");

interface IFactory_7<TSelf> where TSelf : IFactory_7<TSelf> {
    static abstract string Create();
}
class FactoryImpl_7 : IFactory_7<FactoryImpl_7> {
    public static string Create() => "Instance_7";
}
class FactoryCaller_7 {
    public static string Call() => FactoryImpl_7.Create();
}
