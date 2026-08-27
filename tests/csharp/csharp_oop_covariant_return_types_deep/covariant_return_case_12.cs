// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_12

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

var repo = new DerivedRepo_12();
__P((repo.Get() is DerivedEntity_12).ToString());
__Check("True");

class BaseEntity_12 { }
class DerivedEntity_12 : BaseEntity_12 { }
abstract class BaseRepo_12 {
    public abstract BaseEntity_12 Get();
}
class DerivedRepo_12 : BaseRepo_12 {
    public override DerivedEntity_12 Get() => new DerivedEntity_12();
}
