// vybe-test: kotlin/sealed_types/test_simple_sealed_when_exhaustive_without_else
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Result {
            class Ok(val value: Int) : Result()
            class Fail : Result()
        }

        fun describe(result: Result): String {
            return when (result) {
                is Result.Ok -> "ok:" + result.value.toString()
                is Result.Fail -> "fail"
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
            val value = describe(Result.Ok(3))
            val other = describe(Result.Fail())
            __p((value).toString())
            __p((other).toString())
        
__check("ok:3\nfail")
}
