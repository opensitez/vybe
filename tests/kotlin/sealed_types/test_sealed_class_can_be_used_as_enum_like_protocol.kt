// vybe-test: kotlin/sealed_types/test_sealed_class_can_be_used_as_enum_like_protocol
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Status {
            class Success(val message: String) : Status()
            class Failure(val code: Int) : Status()
        }

        fun statusCode(status: Status): Int {
            return when (status) {
                is Status.Success -> 0
                is Status.Failure -> status.code
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
            __p((statusCode(Status.Success("ok"))).toString())
            __p((statusCode(Status.Failure(7))).toString())
        
__check("0\n7")
}
