// vybe-test: kotlin/advanced_features/test_advanced_sealed_with_deeper_matching
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

sealed class Status {
            class Ok(val message: String) : Status()
            class Error(val code: Int) : Status()
        }

        fun summarize(s: Status): String {
            return when (s) {
                is Status.Ok -> s.message
                is Status.Error -> "E" + s.code
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((summarize(Status.Ok("fine"))).toString(), "fine")
            __check((summarize(Status.Error(42))).toString(), "E42")
        }
