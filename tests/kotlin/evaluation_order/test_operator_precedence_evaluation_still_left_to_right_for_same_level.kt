// vybe-test: kotlin/evaluation_order/test_operator_precedence_evaluation_still_left_to_right_for_same_level
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

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
            var order = ""
            fun a(): Int { order += "a"
return 1 }
            fun b(): Int { order += "b"
return 2 }
            fun c(): Int { order += "c"
return 3 }
            val out = a() + b() * c()
            __p((out).toString())
            __p((order).toString())
        
__check("7\nabc")
}
