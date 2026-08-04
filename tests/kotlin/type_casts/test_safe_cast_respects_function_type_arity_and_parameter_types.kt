// vybe-test: kotlin/type_casts/test_safe_cast_respects_function_type_arity_and_parameter_types
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

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
            val handler: Any = { value: Int -> value.toString() }
            val unary = handler as? (Int) -> String
            val binary = handler as? (Int, Int) -> String
            __p((unary != null).toString())
            __p((binary == null).toString())
        
__check("true\ntrue")
}
