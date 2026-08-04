// vybe-test: kotlin/invoke_operator/test_invoke_operator_dispatch_in_subclass
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

open class A {
            open operator fun invoke(v: String): String = "A: " + v
        }
        class B : A() {
            override operator fun invoke(v: String): String = "B: " + v
        }
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
            val a: A = B()
            __p((a("x")).toString())
        
__check("B: x")
}
