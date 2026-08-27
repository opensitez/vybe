// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_2

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

string res = FactoryCaller_2.Call();
__P(res);
__Check("Instance_2");

interface IFactory_2<TSelf> where TSelf : IFactory_2<TSelf> {
    static abstract string Create();
}
class FactoryImpl_2 : IFactory_2<FactoryImpl_2> {
    public static string Create() => "Instance_2";
}
class FactoryCaller_2 {
    public static string Call() => FactoryImpl_2.Create();
}
