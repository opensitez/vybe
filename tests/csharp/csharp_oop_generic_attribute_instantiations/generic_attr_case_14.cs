// vybe-test: csharp/csharp_oop_generic_attribute_instantiations/generic_attr_case_14

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

var attr = (ValidatorAttribute_14<int>)typeof(TargetClass_14).GetCustomAttributes(typeof(ValidatorAttribute_14<int>), false)[0];
__P(attr.TargetType.Name);
__Check("Int32");

[System.AttributeUsage(System.AttributeTargets.Class)]
class ValidatorAttribute_14<T> : System.Attribute {
    public Type TargetType => typeof(T);
}
[ValidatorAttribute_14<int>]
class TargetClass_14 { }
