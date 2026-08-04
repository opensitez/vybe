// vybe-test: kotlin/kotlin_system_utils/test_system_identity_hash_code_repeatability_for_object
// origin: languages/kotlin/tests/kotlin/test_kotlin_system_utils.rs

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
            val value = arrayOf(1, 2, 3)
            val first = System.identityHashCode(value)
            val second = System.identityHashCode(value)
            __p((first != 0).toString())
            __p((second != 0).toString())
            __p((first == second).toString())
        
__check("true\ntrue\ntrue")
}
