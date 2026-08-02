// Vybe test harness — C#.
//
// Real C# alongside harness/go/check.go and harness/js/check.js, the way
// test262's assert.js is JavaScript.
//
// Declared as a LOCAL FUNCTION because the corpus is written as top-level
// statements — there is no class or Main to hang a static method on, and a
// type declaration would have to follow the statements it is used by.
//
// A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
// throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
// says nothing at all.

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}
