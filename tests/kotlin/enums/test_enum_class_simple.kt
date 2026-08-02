// vybe-test: kotlin/enums/test_enum_class_simple
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Direction {
            NORTH, SOUTH, EAST, WEST
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val dir = Direction.NORTH
            __check((dir).toString(), "0")
        }
