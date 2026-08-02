// vybe-test: kotlin/enums/test_enum_when_matching
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class HttpStatus { OK, ERROR, UNKNOWN }

        fun describe(status: HttpStatus): String {
            return when (status) {
                HttpStatus.OK -> "ok"
                HttpStatus.ERROR -> "error"
                HttpStatus.UNKNOWN -> "unknown"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(HttpStatus.ERROR)).toString(), "error")
        }
