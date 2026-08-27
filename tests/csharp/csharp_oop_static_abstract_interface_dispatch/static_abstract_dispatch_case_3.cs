// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_3

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

string res = FactoryCaller_3.Call();
__P(res);
__Check("Instance_3");

interface IFactory_3<TSelf> where TSelf : IFactory_3<TSelf> {
    static abstract string Create();
}
class FactoryImpl_3 : IFactory_3<FactoryImpl_3> {
    public static string Create() => "Instance_3";
}
class FactoryCaller_3 {
    public static string Call() => FactoryImpl_3.Create();
}
