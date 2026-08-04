// vybe-test: kotlin/kotlin_result_runtime/test_result_on_success_and_on_failure_have_side_effects
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

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
            var successSeen = false
            var failureSeen = false
            runCatching { 5 }
                .onSuccess { successSeen = true }
                .onFailure { failureSeen = true }
            runCatching<Int> { throw RuntimeException("fail") }
                .onSuccess { successSeen = true }
                .onFailure { failureSeen = true }
            __p((successSeen).toString())
            __p((failureSeen).toString())
        
__check("true\ntrue")
}
