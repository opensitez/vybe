// vybe-test: kotlin/constructor_chaining/test_constructor_final_read
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class FinalRead {
            val first = 1
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((FinalRead().first).toString(), "1")
        }
