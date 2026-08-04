// vybe-test: kotlin/sealed_types/test_when_on_sealed_interface_in_function_body_with_local_value
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed interface Stage
        class Start : Stage
        class End : Stage

        fun stage_text(stage: Stage): String {
            val value: Stage = stage
            return when (value) {
                is Start -> "start"
                is End -> "end"
            }
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
            val start: Stage = Start()
            val end: Stage = End()
            __p((stage_text(start)).toString())
            __p((stage_text(end)).toString())
        
__check("start\nend")
}
