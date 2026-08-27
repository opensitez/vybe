// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_11

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

var repo = new DerivedRepo_11();
__P((repo.Get() is DerivedEntity_11).ToString());
__Check("True");

class BaseEntity_11 { }
class DerivedEntity_11 : BaseEntity_11 { }
abstract class BaseRepo_11 {
    public abstract BaseEntity_11 Get();
}
class DerivedRepo_11 : BaseRepo_11 {
    public override DerivedEntity_11 Get() => new DerivedEntity_11();
}
