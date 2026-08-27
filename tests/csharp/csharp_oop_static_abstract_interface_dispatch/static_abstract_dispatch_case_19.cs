// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_19

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

string res = FactoryCaller_19.Call();
__P(res);
__Check("Instance_19");

interface IFactory_19<TSelf> where TSelf : IFactory_19<TSelf> {
    static abstract string Create();
}
class FactoryImpl_19 : IFactory_19<FactoryImpl_19> {
    public static string Create() => "Instance_19";
}
class FactoryCaller_19 {
    public static string Call() => FactoryImpl_19.Create();
}
