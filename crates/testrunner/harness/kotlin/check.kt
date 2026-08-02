// Vybe test harness — Kotlin.
//
// Real Kotlin source alongside harness/go/check.go and harness/js/check.js,
// the way test262's assert.js is JavaScript.
//
// A test's verdict is its EXIT CODE. `__check` prints its diagnostic BEFORE
// throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
// says nothing at all.
//
// Kotlin's `println` takes a single argument, so there is no join to mirror —
// unlike Go's Println and JS's console.log, which space-separate.

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}
