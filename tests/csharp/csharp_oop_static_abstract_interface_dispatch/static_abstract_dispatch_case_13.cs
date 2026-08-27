// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_13

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

string res = FactoryCaller_13.Call();
__P(res);
__Check("Instance_13");

interface IFactory_13<TSelf> where TSelf : IFactory_13<TSelf> {
    static abstract string Create();
}
class FactoryImpl_13 : IFactory_13<FactoryImpl_13> {
    public static string Create() => "Instance_13";
}
class FactoryCaller_13 {
    public static string Call() => FactoryImpl_13.Create();
}
