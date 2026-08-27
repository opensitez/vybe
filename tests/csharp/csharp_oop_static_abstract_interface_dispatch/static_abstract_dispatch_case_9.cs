// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_9

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

string res = FactoryCaller_9.Call();
__P(res);
__Check("Instance_9");

interface IFactory_9<TSelf> where TSelf : IFactory_9<TSelf> {
    static abstract string Create();
}
class FactoryImpl_9 : IFactory_9<FactoryImpl_9> {
    public static string Create() => "Instance_9";
}
class FactoryCaller_9 {
    public static string Call() => FactoryImpl_9.Create();
}
