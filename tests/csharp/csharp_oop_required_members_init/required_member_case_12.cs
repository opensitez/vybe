// vybe-test: csharp/csharp_oop_required_members_init/required_member_case_12

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

var p = new Person_12 { Name = "User_12" };
__P(p.Name);
__Check("User_12");

class Person_12 {
    public required string Name { get; init; }
}
