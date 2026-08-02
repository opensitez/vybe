// vybe-test: kotlin/when_subjects/test_when_boolean
// origin: languages/kotlin/tests/kotlin/test_when_subjects.rs

fun resolve(x: Boolean): String = when (x) {
            true -> "yes"
            false -> "no"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((resolve(true)).toString(), "yes")
            __check((resolve(false)).toString(), "no")
        }
