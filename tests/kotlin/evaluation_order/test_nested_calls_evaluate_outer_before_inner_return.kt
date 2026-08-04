// vybe-test: kotlin/evaluation_order/test_nested_calls_evaluate_outer_before_inner_return
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun combine(x: Int, y: Int): Int = x * 10 + y
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
            fun f(): Int { order += "f"
return 1 }
            fun g(x: Int): Int { order += "g"
return x + 1 }
            fun h(x: Int): Int { order += "h"
return x + 2 }
            val out = combine(g(f()), h(g(3)))
            __p((out).toString())
            __p((order).toString())
        
__check("26\nfghg")
}
