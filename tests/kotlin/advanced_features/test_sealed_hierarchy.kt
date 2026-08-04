// vybe-test: kotlin/advanced_features/test_sealed_hierarchy
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

sealed class Result {
            class Ok(val value: Int) : Result()
            class Error(val message: String) : Result()
        }

        fun format(result: Result): String {
            return when (result) {
                is Result.Ok -> "ok:" + (result.value)
                is Result.Error -> "error:" + (result.message)
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
            val good = Result.Ok(7)
            val bad = Result.Error("bad")
            __p((format(good)).toString())
            __p((format(bad)).toString())
        
__check("ok:7\nerror:bad")
}
