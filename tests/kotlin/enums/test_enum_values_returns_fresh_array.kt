// vybe-test: kotlin/enums/test_enum_values_returns_fresh_array
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Color { RED, GREEN, BLUE }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Color.values()
            val second = Color.values()
            first[0] = Color.GREEN
            __check((second[0] == Color.RED).toString(), "true")
            __check((first[0] == Color.GREEN).toString(), "true")
        }
