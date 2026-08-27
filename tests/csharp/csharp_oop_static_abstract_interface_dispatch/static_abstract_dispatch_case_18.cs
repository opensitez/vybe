// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_18

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

string res = FactoryCaller_18.Call();
__P(res);
__Check("Instance_18");

interface IFactory_18<TSelf> where TSelf : IFactory_18<TSelf> {
    static abstract string Create();
}
class FactoryImpl_18 : IFactory_18<FactoryImpl_18> {
    public static string Create() => "Instance_18";
}
class FactoryCaller_18 {
    public static string Call() => FactoryImpl_18.Create();
}
