// vybe-test: kotlin/kotlin_unit_type/test_unit_return_is_executable
// origin: languages/kotlin/tests/kotlin/test_kotlin_unit_type.rs

var marker = 0

        fun stamp(value: Int): Unit {
            marker = value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            stamp(7)
            __check((marker).toString(), "7")
        }
