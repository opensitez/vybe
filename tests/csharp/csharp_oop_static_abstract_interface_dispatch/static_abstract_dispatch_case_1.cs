// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_1

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

string res = FactoryCaller_1.Call();
__P(res);
__Check("Instance_1");

interface IFactory_1<TSelf> where TSelf : IFactory_1<TSelf> {
    static abstract string Create();
}
class FactoryImpl_1 : IFactory_1<FactoryImpl_1> {
    public static string Create() => "Instance_1";
}
class FactoryCaller_1 {
    public static string Call() => FactoryImpl_1.Create();
}
