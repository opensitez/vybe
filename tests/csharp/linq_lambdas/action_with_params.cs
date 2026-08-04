// vybe-test: csharp/linq_lambdas/action_with_params
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

Action<string, int> describe = (name, age) => {
    __P((name + " is " + age).ToString());
};
describe("Alice", 30);
describe("Bob", 25);
__Check("Alice is 30\nBob is 25");
