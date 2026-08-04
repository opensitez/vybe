// vybe-test: kotlin/sealed_types/test_non_exhaustive_when_still_requires_else_for_non_sealed
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Alpha {
            class A : Alpha()
        }

        open class Beta

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
            val alpha: Alpha = Alpha.A()
            __p((when (alpha) {
                is Alpha.A -> 1
            }).toString())
            val beta = Beta()
            __p((when (beta is Beta) {
                true -> 2
                false -> 3
            }).toString())
        
__check("1\n2")
}
