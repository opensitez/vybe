// vybe-test: kotlin/kotlin_result_api/test_result_exception_message_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

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
            val value = runCatching<Int> { 1 / 0 }
                .map { it + 1 }
                .recover { 0 }
            val thrown = runCatching<Int> { 1 / 0 }
                .exceptionOrNull()
            __p((value).toString())
            __p((thrown?.let { it.message }).toString())
        
__check("0\n/ by zero")
}
