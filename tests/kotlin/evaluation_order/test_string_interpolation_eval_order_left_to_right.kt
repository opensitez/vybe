// vybe-test: kotlin/evaluation_order/test_string_interpolation_eval_order_left_to_right
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
            fun left(): String { order += "L"
return "x" }
            fun right(): String { order += "R"
return "y" }
            val result = "${left()}${right()}"
            __p((result).toString())
            __p((order).toString())
        
__check("xy\nLR")
}
