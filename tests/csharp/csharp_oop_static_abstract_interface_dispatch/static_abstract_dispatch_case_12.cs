// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_12

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

string res = FactoryCaller_12.Call();
__P(res);
__Check("Instance_12");

interface IFactory_12<TSelf> where TSelf : IFactory_12<TSelf> {
    static abstract string Create();
}
class FactoryImpl_12 : IFactory_12<FactoryImpl_12> {
    public static string Create() => "Instance_12";
}
class FactoryCaller_12 {
    public static string Call() => FactoryImpl_12.Create();
}
