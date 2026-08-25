// vybe-test: kotlin/local_classes/test_local_class_private_constructor_not_exposed
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

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

// Damaged spelling repaired: the class lived inside `fun main()` with a
// companion, and kotlinc 2.4.10 rejects that — "modifier 'companion' is not
// applicable inside 'local class'". Hoisted to top level, which keeps the
// test's point (a private constructor is reachable only through the factory);
// measured under kotlinc 2.4.10 it prints the expected 2.
class C private constructor(val v: Int) {
                companion object {
                    fun make(v: Int) = C(v)
                }
            }

fun main() {
            __p((C.make(2).v).toString())

__check("2")
}
