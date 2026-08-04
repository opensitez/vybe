// vybe-test: kotlin/exceptions/test_run_catching_map_and_recover_flow
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

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
            val value = runCatching { "k".toInt() }
                .map { it + 1 }
                .onFailure { __p(("fail").toString()) }
                .recover { 9 }

            __p((value.getOrNull()).toString())
            __p((value.isFailure).toString())
        
__check("fail\n9\nfalse")
}
