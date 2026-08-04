// Vybe test harness — C#.
//
// Real C# alongside harness/go/check.go and harness/js/check.js, the way
// test262's assert.js is JavaScript.
//
// Declared as LOCAL FUNCTIONS over a local `__buf` because the corpus is
// written as top-level statements — there is no class or Main to hang a static
// method on, and a type declaration would have to follow the statements it is
// used by. A local function captures `__buf` by reference, so the appends are
// visible to `__Check` at the end of the file.
//
// A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
// throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
// says nothing at all.
//
// Output is COLLECTED, not paired. The emitter rewrites every
// `Console.WriteLine(x)` into `__P((x).ToString())` and compares the whole
// output once at the end. Pairing the i-th print with the i-th expected line
// cannot assert anything about a loop, and loops alone were 706 of C#'s 7,622
// cases.
//
// Rendering happens at the CALL SITE, and with `.ToString()` specifically.
// Measured against vybex: `WriteLine(b)` and `b.ToString()` both give `True`,
// while `Convert.ToString(b)` gives `true` and `"" + b` gives `0` — the wrong
// one cost 3,323 false failures. (Real C# gives `True` for all three, so those
// two are Vybe bugs in their own right.)

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
