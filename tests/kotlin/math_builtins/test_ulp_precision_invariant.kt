// vybe-test: kotlin/math_builtins/test_ulp_precision_invariant
// origin: languages/kotlin/tests/kotlin/test_math_builtins.rs

// `kotlin.math.ulp(v)` and `kotlin.math.nextAfter(a, b)` do not exist —
// `kotlinc` rejects both with "unresolved reference". `ulp` is an extension
// PROPERTY (`v.ulp`) and the function is `nextTowards`. Repaired against
// kotlinc, which compiles this form clean and prints true/true/true/true.
import kotlin.math.ulp
import kotlin.math.nextTowards

var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = 1.0
            val step = v.ulp
            val nearOne = 1.0 + step
            val isAdjacent = 1.0.nextTowards(Double.POSITIVE_INFINITY) == nearOne
            __p((nearOne > v).toString())
            __p((step > 0.0).toString())
            __p((isAdjacent).toString())
            __p((kotlin.math.abs(step - nearOne.ulp) < 1e-20).toString())
        
__check("true\ntrue\ntrue\ntrue")
}
