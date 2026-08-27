// vybe-test: csharp/csharp_oop_static_abstract_interface_dispatch/static_abstract_dispatch_case_17

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

string res = FactoryCaller_17.Call();
__P(res);
__Check("Instance_17");

interface IFactory_17<TSelf> where TSelf : IFactory_17<TSelf> {
    static abstract string Create();
}
class FactoryImpl_17 : IFactory_17<FactoryImpl_17> {
    public static string Create() => "Instance_17";
}
class FactoryCaller_17 {
    public static string Call() => FactoryImpl_17.Create();
}
