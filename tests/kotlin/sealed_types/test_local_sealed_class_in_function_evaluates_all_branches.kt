// vybe-test: kotlin/sealed_types/test_local_sealed_class_in_function_evaluates_all_branches
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

fun classify(flag: Boolean): String {
            sealed class LocalResult {
                class Yes(val label: String) : LocalResult()
                class No : LocalResult()
            }

            val result: LocalResult = if (flag) LocalResult.Yes("ok") else LocalResult.No()
            return when (result) {
                is LocalResult.Yes -> result.label
                is LocalResult.No -> "no"
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
            __p((classify(true)).toString())
            __p((classify(false)).toString())
        
__check("ok\nno")
}
